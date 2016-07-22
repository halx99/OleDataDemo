/*
This file is part of the OleDataDemo example
Drag & drop images and drop descriptions for MFC applications
hosted at CodeProject, <http://www.codeproject.com>.

Copyright (C) 2015 Jochen Arndt

This software is licensed under the terms of the The Code Project Open License (CPOL).
You should have received a copy of the license along with this program.
If not, see <http://www.codeproject.com/info/cpol10.aspx>.

To contact the author use the forum of the corresponding CodeProject article.
The author's articles are listed at <http://www.codeproject.com/Articles/Jochen-Arndt#articles>.

*/

// InfoList.cpp : Show clipboard and drag & drop data object content.
//

#include "stdafx.h"
#include "OleDataDemo.h"
#include "OleDataDemoDlg.h"
#include "OleDropTargetEx.h"
#include "InfoList.h"
#include ".\infolist.h"

#include <richedit.h>	// RTF related format names
#include <shlobj.h>		// Shell related format names
//#include <urlmon.h>		// MIME related format names
#include <hlink.h>		// IHlink interface, CFSTR_HYPERLINK
#include <mmreg.h>		// Wave format

#pragma comment(lib, "hlink")

#define EXCERPT_LEN			100
#define ISTREAM_BUF_SIZE	2048					// max. number of bytes to read from IStream

// Headers for some binary file formats

#pragma pack(1)
// ZIP local file header
typedef struct { 
	DWORD dwSig;					// local file header signature     4 bytes  (0x04034b50)
	WORD wMinVer;					// version needed to extract       2 bytes
	WORD wBitFlag;					// general purpose bit flag        2 bytes
	WORD wCompression;				// compression method              2 bytes
	WORD wModTime;					// last mod file time              2 bytes
	WORD wModDate;					// last mod file date              2 bytes
	DWORD dwCRC32;					// crc-32                          4 bytes
	DWORD dwCompressedSize;			// compressed size                 4 bytes
	DWORD dwUncompressedSize;		// uncompressed size               4 bytes
	WORD wFileNameLength;			// file name length                2 bytes
	WORD wExtraFieldLength;			// extra field length              2 bytes
	char lpszFileName[1];			// file name (variable size)
									// extra field (variable size)
} ZipHdr_t;

// RIFF header
typedef struct { 
	char chunkID[4];				// "RIFF"
	DWORD chunkSize;				// file size - 8
	char type[4];					// e.g. "WAVE"
	char fmt[4];					// "fmt "
	DWORD fmtSize;					// size of following format section
	WAVEFORMAT wf;					// WAVE only
} RiffHdr_t;
#pragma pack()

// LARGE_INTEGER null value used with IStream::Seek().
const LARGE_INTEGER CDataInfo::m_LargeNull = { 0, 0 };

const unsigned CInfoList::m_nDefaultColumns = 
	(1 << FormatColumn) + (1 << TymedColumn) + 
	(1 << SizeColumn) + (1 << TypeColumn) + (1 << ContentColumn);

// CInfoList

IMPLEMENT_DYNAMIC(CInfoList, CListCtrl)
CInfoList::CInfoList()
{
	m_bCurrentLocale = false;
	m_bAllTymeds = false;
	m_nVisibleColumns = m_nAutoColumns = m_nDefaultColumns;
	m_hWndChain = NULL;
}

CInfoList::~CInfoList()
{
}


BEGIN_MESSAGE_MAP(CInfoList, CListCtrl)
	ON_WM_DRAWCLIPBOARD()
	ON_WM_CHANGECBCHAIN()
	ON_WM_DESTROY()
END_MESSAGE_MAP()

//
// Drag callback functions
//

DROPEFFECT CInfoList::OnDragEnter(CWnd *pWnd, COleDataObject* /*pDataObject*/, DWORD /*dwKeyState*/, CPoint point)
{
	ASSERT(pWnd->IsKindOf(RUNTIME_CLASS(CInfoList)));
	CInfoList *pThis = static_cast<CInfoList*>(pWnd);
	return pThis->GetDropEffect(point);
}

// Show special cursor and description text
DROPEFFECT CInfoList::OnDragOver(CWnd *pWnd, COleDataObject* /*pDataObject*/, DWORD /*dwKeyState*/, CPoint point)
{
	ASSERT(pWnd->IsKindOf(RUNTIME_CLASS(CInfoList)));
	CInfoList *pThis = static_cast<CInfoList*>(pWnd);
	return pThis->GetDropEffect(point);
}

// Fill the list with data object info upon dropping.
DROPEFFECT CInfoList::OnDropEx(CWnd *pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT /*dropList*/, CPoint /*point*/)
{
	ASSERT(pWnd->IsKindOf(RUNTIME_CLASS(CInfoList)));
	if (dropDefault)
	{
		CInfoList *pThis = static_cast<CInfoList*>(pWnd);
		pThis->UpdateFromData(pDataObject);
	}
	return DROPEFFECT_NONE;							// May also return DROPEFFECT_COPY here.
}

// Set drop description and get drop effect for OnDragEnter() and OnDragOver().
DROPEFFECT CInfoList::GetDropEffect(CPoint point)
{
	DROPEFFECT dwEffect = DROPEFFECT_NONE;
	CRect rc;
	GetClientRect(rc);
	if (rc.PtInRect(point))							// Exclude scroll bars
	{
		// Example of a target specific drop description.
		// Show DROPIMAGE_LABEL icon and specified text if source has enabled display of description text.
		// The DropDescription data object is created if necessary.
		// Uses the message text "Update %1" and the insert text "Info List" 
		//  which has been specified in Init().
		m_DropTarget.SetDropDescription(DROPIMAGE_LABEL, NULL, true);
		// Must return a drop effect. Otherwise old style cursor would be the stop symbol.
		dwEffect = DROPEFFECT_COPY;
	}
	else
	{
		// Clear drop description so that the default cursor and text are shown.
		m_DropTarget.ClearDropDescription();
	}
	return dwEffect;
}

// WM_DRAWCLIPBORAD: Clipboard content has changed.
void CInfoList::OnDrawClipboard()
{
	CListCtrl::OnDrawClipboard();

	UpdateFromClipboard();

	// Pass message to next window in chain.
	if (m_hWndChain)
		::SendMessage(m_hWndChain, WM_DRAWCLIPBOARD, 0, 0);
}

// WM_CHANGECBCHAIN: Clipboard viewer chain update
void CInfoList::OnChangeCbChain(HWND hWndRemove, HWND hWndAfter)
{
	CListCtrl::OnChangeCbChain(hWndRemove, hWndAfter);
	// Next window in chain is removed.
	// So update our member var with the new next window.
	if (m_hWndChain == hWndRemove)
		m_hWndChain = hWndAfter;
	// Pass message over to next viewer
	else if (m_hWndChain)
		::SendMessage(m_hWndChain, WM_CHANGECBCHAIN, (WPARAM)hWndRemove, (LPARAM)hWndAfter);
}

void CInfoList::OnDestroy()
{
	// Remove this window from the chain of clipboard viewers.
	if (m_hWndChain)
		ChangeClipboardChain(m_hWndChain);
	CListCtrl::OnDestroy();
}

void CInfoList::Init(COleDataDemoDlg *p /*= NULL*/, bool bCurrentLocale /*= false*/)
{
	m_pParentDlg = p;
	m_bCurrentLocale = bCurrentLocale;

	// m_pDropTarget is CWnd member that is set to the address of the COleDropTarget class
	// when registering. If it is not NULL, this class has been already registered.
    if (NULL == m_pDropTarget)
    {
		// Enable new styles for list controls.
		if (::IsAppThemed() && (::GetThemeAppProperties() & STAP_ALLOW_CONTROLS))
			VERIFY(S_OK == ::SetWindowTheme(m_hWnd, L"Explorer", NULL));

		// Set callback functions.
		m_DropTarget.SetOnDragEnter(OnDragEnter);
		m_DropTarget.SetOnDragOver(OnDragOver);
		m_DropTarget.SetOnDropEx(OnDropEx);
		// Show special cursor and text. Requires calling of m_DropTarget.SetDropDescription().
		// Using an insert text is not really necessary but it is shown in a different color.
		m_DropTarget.SetDropDescriptionText(DROPIMAGE_LABEL, NULL, L"Update %1");
		m_DropTarget.SetDropInsertText(L"Info List");
		// Register as drop target.
		VERIFY(m_DropTarget.Register(this));
	}

	if (0 == GetHeaderCtrl()->GetItemCount())
	{
		InsertColumn(InfoColumns::FormatColumn, _T("Format"), LVCFMT_LEFT, 100, InfoColumns::FormatColumn);
		InsertColumn(InfoColumns::TymedColumn, _T("Storage Medium"), LVCFMT_LEFT, 100, InfoColumns::TymedColumn);
		InsertColumn(InfoColumns::UsedTymedColumn, _T("Used Medium"), LVCFMT_LEFT, 100, InfoColumns::UsedTymedColumn);
		InsertColumn(InfoColumns::DeviceColumn, _T("Device"), LVCFMT_LEFT, 60, InfoColumns::DeviceColumn);
		InsertColumn(InfoColumns::AspectColumn, _T("Aspect"), LVCFMT_LEFT, 60, InfoColumns::AspectColumn);
		InsertColumn(InfoColumns::IndexColumn, _T("Index"), LVCFMT_RIGHT, 100, InfoColumns::IndexColumn);
		InsertColumn(InfoColumns::ReleaseColumn, _T("Release"), LVCFMT_LEFT, 100, InfoColumns::ReleaseColumn);
		InsertColumn(InfoColumns::SizeColumn, _T("Size"), LVCFMT_RIGHT, 60, InfoColumns::SizeColumn);
		InsertColumn(InfoColumns::TypeColumn, _T("Data Type"), LVCFMT_LEFT, 60, InfoColumns::TypeColumn);
		InsertColumn(InfoColumns::ContentColumn, _T("Content"), LVCFMT_LEFT, 200, InfoColumns::ContentColumn);

		// Add this window to the list of clipboard viewers.
		// This will also update the list with the current clipboard Info.
		if (NULL == m_hWndChain)
			m_hWndChain = SetClipboardViewer();
	}
}

// Fill the list with data on the clipboard.
void CInfoList::UpdateFromClipboard()
{
	COleDataObject OleDataObject;
	OleDataObject.AttachClipboard();
	UpdateFromData(&OleDataObject);
	OleDataObject.Release();
}

// Fill the list with data from COleDataObject.
void CInfoList::UpdateFromData(COleDataObject* pDataObj)
{
	ASSERT(pDataObj);

	// Update title of info list in main dialog.
	// NOTE: pDataObj->m_bClipboard is cleared when calling the first data function!
	if (m_pParentDlg)
		m_pParentDlg->SetInfoTitle(pDataObj->m_bClipboard ? _T("Clipboard Content") : _T("Drag Data"));
	SetRedraw(0);									// Disable redrawing while updating the list
	DeleteAllItems();

	m_nAutoColumns = m_nDefaultColumns;
	CDataInfo Info(&m_DropTarget, m_bCurrentLocale);
	FORMATETC FormatEtc;
	pDataObj->BeginEnumFormats();
	while (pDataObj->GetNextFormat(&FormatEtc))
	{
		// Note that FormatEtc.tymed can contain multiple storage types.
		// So we must select each type or let the GetData() implementation
		//  decide which type to use.
		if (m_bAllTymeds)
		{
			DWORD tymedIn = FormatEtc.tymed;
			DWORD tymedOut = 1;
			while (tymedIn)
			{
				if (tymedIn & 1)
				{
					m_nAutoColumns |= Info.GetInfo(tymedOut, FormatEtc, pDataObj);
					AddItem(Info);
				}
				tymedIn >>= 1;
				tymedOut <<= 1;
			}
		}
		// This will retrieve data using a tymed selected by the GetData() implementation.
		// E.g. with HGLOBAL and ISTREAM, HGLOBAL may be used.
		else
		{
			m_nAutoColumns |= Info.GetInfo(TYMED_NULL, FormatEtc, pDataObj);
			AddItem(Info);
		}
	}
	SizeColumns(m_nAutoColumns);
	SetRedraw(1);
}

// Hide invisible columns and autosize the others.
void CInfoList::SizeColumns(unsigned nMask /*= 0*/)
{
	if (nMask)
		m_nVisibleColumns = nMask;
	int nColumns = GetHeaderCtrl()->GetItemCount();
	for (int i = 0; i < nColumns; i++)
		SetColumnWidth(i, (m_nVisibleColumns & (1 << i)) ? LVSCW_AUTOSIZE_USEHEADER : 0);
}

int CInfoList::AddItem(const CDataInfo& Info)
{
	CString str;

	// Clipboard format
	str.Format(_T("%s (%#X)"), Info.GetFormatName(), Info.GetFormat());
	int nItem = InsertItem(INT_MAX, str.GetString());
//	SetItemData(nItem, Info.GetFormat());

	// Type of storage medium (may be multiple types).
	str.Format(_T("%s (%#X)"), Info.GetTymedString(Info.GetTymed()).GetString(), Info.GetTymed());
	SetItemText(nItem, InfoColumns::TymedColumn, str.GetString());

	// The used tymed may be one that is not set in the FORMATETC structure! 
	str.Format(_T("%s (%#X)"), Info.GetTymedString(Info.GetUsedTymed()).GetString(), Info.GetUsedTymed());
	SetItemText(nItem, InfoColumns::UsedTymedColumn, str.GetString());

	// Device
	SetItemText(nItem, InfoColumns::DeviceColumn, Info.GetDeviceString().GetString());

	// Aspect
	SetItemText(nItem, InfoColumns::AspectColumn, Info.GetAspectString());

	// Index
	str.Format(_T("%d"), Info.GetIndex());
	SetItemText(nItem, InfoColumns::IndexColumn, str.GetString());

	// Released by Provider or Receiver
	SetItemText(nItem, InfoColumns::ReleaseColumn, Info.GetUnkForRelease() ? _T("Provider") : _T("Receiver"));

	// Storage size
	SetItemText(nItem, InfoColumns::SizeColumn, Info.PrintSize(Info.GetSize()));

	// Data type or text encoding
	if (Info.HasDataType())
		SetItemText(nItem, InfoColumns::TypeColumn, Info.GetDataType());

	// Content information
	if (Info.HasContent())
		SetItemText(nItem, InfoColumns::ContentColumn, Info.GetContent());
	return nItem;
}

// Helper class to get and store data object information

CDataInfo::CDataInfo(const COleDropTargetEx *pDropTarget, bool bCurrentLocale /*= false*/)
{
	m_bBufferAllocated = false;
	m_pDropTarget = pDropTarget;
	m_nCP = (UINT)-1;
	m_nSize = m_nBufSize = 0;
	m_lpData = NULL;
	_tcscpy_s(m_lpszThSep, ARRAYSIZE(m_lpszThSep) , _T(","));
	// Thousands separator character.
	// Optional use locale of current thread (or pass LOCALE_SYSTEM_DEFAULT or LOCALE_USER_DEFAULT).
	if (bCurrentLocale)
		::GetLocaleInfo(::GetThreadLocale(), LOCALE_STHOUSAND, m_lpszThSep, ARRAYSIZE(m_lpszThSep));
}

CDataInfo::~CDataInfo()
{
	FreeBuffer();
}

// Allocate buffer for non HGLOBAL data.
LPVOID CDataInfo::AllocBuffer(size_t nSize)
{
	FreeBuffer();
	if (nSize)
	{
		m_lpData = new BYTE[nSize];
		m_nBufSize = nSize;
		m_bBufferAllocated = true;
	}
	return m_lpData;
}

void CDataInfo::FreeBuffer()
{
	if (m_bBufferAllocated)
		delete [] m_lpData;
	m_lpData = NULL;
	m_nBufSize = 0;
	m_bBufferAllocated = false;
}

// Get info about data object.
unsigned CDataInfo::GetInfo(DWORD tymed, const FORMATETC& FormatEtcIn, COleDataObject* pDataObj)
{
	ASSERT(pDataObj);

	unsigned nColumnsMask = 0;
	m_nCP = (UINT)-1;								// code page with text data 
	m_nSize = 0;									// storage size
	m_nUsedTymed = TYMED_NULL;						// the tymed finally used to retrive the data
	m_pUnkForRelease = NULL;						// StgMedium.pUnkForRelease (StgMedium will be released here)
	m_strContent =									// assign empty string here; CString::Empty()
		m_strDataType = m_strText = _T("");			//  would release the memory resulting in multiple (re-)allocations
	m_FormatEtc = FormatEtcIn;						//
	if (TYMED_NULL != tymed)						// set single tymed
		m_FormatEtc.tymed = tymed;	
	FreeBuffer();

	SetFormatName(m_FormatEtc.cfFormat);

	// If there are non-default settings, show the corresponding columns
	if (m_FormatEtc.lindex != -1)					// Index used (e.g. with file contents)
		nColumnsMask |= 1 << CInfoList::IndexColumn;
	if (m_FormatEtc.dwAspect != DVASPECT_CONTENT)
		nColumnsMask |= 1 << CInfoList::AspectColumn;
	if (m_FormatEtc.ptd != NULL)
		nColumnsMask |= 1 << CInfoList::DeviceColumn;

	// GetData() may fail with zero size objects.
	STGMEDIUM StgMedium;
	if (!pDataObj->GetData(m_FormatEtc.cfFormat, &StgMedium, &m_FormatEtc))
	{
		nColumnsMask |= 1 << CInfoList::UsedTymedColumn;
		return nColumnsMask;
	}
	// Store medium data for later use. StgMedium is released below.
	// Note that the tymed finally used to retrieve the data may be any
	// (even one that is not indicated by the FORMATECT struct).
	m_nUsedTymed = StgMedium.tymed;
	m_pUnkForRelease = StgMedium.pUnkForRelease;

	if (StgMedium.pUnkForRelease != NULL)
		nColumnsMask |= 1 << CInfoList::ReleaseColumn;
	if (!(m_FormatEtc.tymed & StgMedium.tymed) ||		// The used tymed is not part of the listed ones or contains multiple tymeds
		CDragDropHelper::PopCount(m_FormatEtc.tymed) > 1)
		nColumnsMask |= 1 << CInfoList::UsedTymedColumn;

	// Get data pointer and size according to the storage type.
	switch (StgMedium.tymed)
	{
	case TYMED_HGLOBAL :							// Global data
	case TYMED_MFPICT :								//  Just get the memory pointer
		m_nSize = m_nBufSize = ::GlobalSize(StgMedium.hGlobal);
		m_lpData = ::GlobalLock(StgMedium.hGlobal);
		break;
	case TYMED_FILE :								// File: Full path stored as OLESTR (allocated by CoTaskMemAlloc())
		{											//  Read some bytes from the file into allocated buffer
			HANDLE hFile = ::CreateFile(OLE2T(StgMedium.lpszFileName), GENERIC_READ, 0, NULL, OPEN_EXISTING, 0, NULL);
			if (INVALID_HANDLE_VALUE != hFile)
			{
				BY_HANDLE_FILE_INFORMATION FileInfo;
				if (::GetFileInformationByHandle(hFile, &FileInfo))
				{
					m_nSize = FileInfo.nFileSizeLow + ((unsigned __int64)FileInfo.nFileSizeHigh << 32ULL);
					AllocBuffer(min(m_nSize, ISTREAM_BUF_SIZE));
					DWORD dwRead = 0;
					if (!::ReadFile(hFile, m_lpData, static_cast<DWORD>(m_nBufSize), &dwRead, NULL))
						FreeBuffer();
					else
						m_nBufSize = dwRead;
				}
				::CloseHandle(hFile);
			}
		}
		break;
	case TYMED_ISTREAM :							// IStream
		{											//  Read some bytes from begin of stream into allocated buffer
			STATSTG st;
			if (SUCCEEDED(StgMedium.pstm->Stat(&st, 0)))
			{
				m_nSize = static_cast<size_t>(st.cbSize.QuadPart);
				// If the storage object has a name, print it to m_strContent and free the name buffer.
				AppendOleStr(st.pwcsName, _T("Stream name"));
				// Read some byte from the stream after rewinding.
				// Actual position is usually at the end so that reading would fail.
				if (SUCCEEDED(StgMedium.pstm->Seek(m_LargeNull, STREAM_SEEK_SET, NULL)))
				{
					AllocBuffer(min(m_nSize, ISTREAM_BUF_SIZE));
					ULONG nRead = 0;
					if (FAILED(StgMedium.pstm->Read(m_lpData, static_cast<ULONG>(m_nBufSize), &nRead)))
						FreeBuffer();
					else
						m_nBufSize = nRead;
				}
			}
		}
		break;
	case TYMED_ISTORAGE :							// IStorage
		{
			STATSTG st;
			if (SUCCEEDED(StgMedium.pstg->Stat(&st, 0)))
			{
				m_nSize = static_cast<size_t>(st.cbSize.QuadPart);
				// If the storage object has a name, print it to m_strContent and free the name buffer.
				AppendOleStr(st.pwcsName, _T("Storage name"));
				// If the storage object has a class identifier, print it to m_strContent.
				AppendCLSID(st.clsid, _T("CLSID"));
			}
		}
		break;
	case TYMED_GDI :								// GDI objects: 
		AllocBuffer(sizeof(DIBSECTION));			//  Call GetObject using allocated buffer with max. object size.
		m_nBufSize = ::GetObject(StgMedium.hBitmap, sizeof(DIBSECTION), m_lpData);
		break;
	case TYMED_ENHMF :								// Enhanced meta file
		{											//  Get header into allocated buffer
			LPENHMETAHEADER pHdr = static_cast<LPENHMETAHEADER>(AllocBuffer(sizeof(ENHMETAHEADER)));
			if (::GetEnhMetaFileHeader(StgMedium.hEnhMetaFile, sizeof(ENHMETAHEADER), pHdr))
			{
				// Don't know if this is correct.
				m_nSize = pHdr->nSize + pHdr->nDescription + pHdr->nPalEntries * sizeof(PALETTEENTRY) + pHdr->nBytes;
			}
			else
				FreeBuffer();
		}
		break;
	}

	// Get information for the specific clipboard format
	if (m_FormatEtc.cfFormat >= 0xC000)
		GetCustomFormatInfo(pDataObj, StgMedium);
	else
		GetStandardFormatInfo(pDataObj, StgMedium);

	// If type of data has not been identified so far
	if (m_strDataType.IsEmpty() && (UINT)-1 == m_nCP && m_nBufSize && m_lpData)
	{
		if (!CheckBinaryFormats())					// Check if data is a known binary format.
		{
			CheckForString();						// Check if data is some kind of text
			if ((UINT)-1 == m_nCP)					// Not some kind of text
				GetBinaryContent();					// Print begin of data as hex bytes
		}
	}
	GetTextInfo();									// Get text encoding for text data

	// Cleanup
	if (TYMED_HGLOBAL == StgMedium.tymed || TYMED_MFPICT == StgMedium.tymed)
		::GlobalUnlock(StgMedium.hGlobal);
	::ReleaseStgMedium(&StgMedium);
	FreeBuffer();
	return nColumnsMask;
}

// Get information for standard clipboard formats
void CDataInfo::GetStandardFormatInfo(COleDataObject* pDataObj, STGMEDIUM& StgMedium)
{
	switch (m_FormatEtc.cfFormat)
	{
	case CF_TEXT :									// get code page from CF_LOCALE or use system default LCID
	case CF_OEMTEXT :
	case CF_DSPTEXT :
		m_nCP = m_pDropTarget->GetCodePage(pDataObj, m_FormatEtc.cfFormat);
		break;
	case CF_BITMAP :
	case CF_DSPBITMAP :
		// m_lpData has been filled by calling GetObject() and m_nBufSize has been set to the return value
		m_strDataType = _T("HBITMAP");
		if (HasData(sizeof(BITMAP)))
		{
			if (sizeof(DIBSECTION) == m_nBufSize)
				m_strDataType = _T("DIBSECTION");
			const BITMAP *pBM = static_cast<const BITMAP*>(m_lpData);
			size_t nImageSize = pBM->bmWidthBytes * pBM->bmHeight;
			// Global memory used by bitmaps.
			// There are 10 + 18 additional bytes stored in the GDI heap.
			// See http://msdn.microsoft.com/en-us/library/ms969928.aspx
			m_nSize = 32 + nImageSize;
			SeparateContent();
			m_strContent.AppendFormat(_T("Image size: %s, width: %d, height: %d, color depth: %hu BPP"), 
				PrintSize(nImageSize), pBM->bmWidth, pBM->bmHeight, pBM->bmBitsPixel);
		}
		break;
	case CF_METAFILEPICT :
	case CF_DSPMETAFILEPICT :
		m_strDataType = _T("METAFILEPICT");
		if (HasData(sizeof(METAFILEPICT)) && StgMedium.tymed != TYMED_GDI)
		{
			// CF_METAFILEPICT uses a HGLOBAL to store a METAFILEPICT structure.
			// The hMF member of this structure is the handle to the meta file.
			const METAFILEPICT *lpMFP = static_cast<const METAFILEPICT*>(m_lpData);
			size_t nFileSize = ::GetMetaFileBitsEx(lpMFP->hMF, 0, NULL);
//			m_nSize += nFileSize;
			SeparateContent();
			m_strContent.AppendFormat(_T("File size: %s, xExt: %d, yExt: %d, mapping mode: "), 
				PrintSize(nFileSize), lpMFP->xExt, lpMFP->yExt);
			if (lpMFP->mm >= MM_TEXT && lpMFP->mm <= MM_ANISOTROPIC)
			{
				LPCTSTR lpszMapMode[] = { _T("TEXT"), _T("LOMETRIC"), _T("HIMETRIC"),
					_T("LOENGLISH"), _T("HIENGLISH"), _T("TWIPS"), _T("ISOTROPIC"), _T("ANISOTROPIC") };
				m_strContent.AppendFormat(_T("MM_%s"), lpszMapMode[lpMFP->mm - 1]);
			}
			else
				m_strContent.AppendFormat(_T("%d"), lpMFP->mm);
		}
		break;
	case CF_SYLK :									// ANSI encoded text
		m_nCP = CP_ACP;
		break;
	case CF_DIF :									// ASCII text
		m_nCP = 20127;
		break;
//	case CF_TIFF :									// Should be identified later by CheckBinaryFormats()
//		break;
	case CF_PALETTE :
		// m_lpData has been filled by calling GetObject() and m_nBufSize has been set to the return value
		m_strDataType = _T("HPALETTE");
		if (HasData(sizeof(WORD)))
		{
			const WORD *pWord = static_cast<const WORD*>(m_lpData);
			// Global memory used by palettes.
			// There are 10 + 10 additional bytes stored in the GDI heap.
			// See http://msdn.microsoft.com/en-us/library/ms969928.aspx
			m_nSize = 4 + 10 * *pWord;
			SeparateContent();
			m_strContent.AppendFormat(_T("%hu entries"), *pWord);
		}
		break;
//	case CF_PENDATA :								// Should be no longer present nowadays
//		break;
//	case CF_RIFF :									// Should be identified later by CheckBinaryFormats()
//		break;
//	case CF_WAVE :									// Should be identified later by CheckBinaryFormats()
//		break;
	case CF_UNICODETEXT :							// Unicode text
		m_nCP = 1200;
		break;
	case CF_ENHMETAFILE :
	case CF_DSPENHMETAFILE :
		m_strDataType = _T("HENHMETAFILE");
		if (HasData(sizeof(ENHMETAHEADER)))
		{
			const ENHMETAHEADER *pHdr = static_cast<const ENHMETAHEADER *>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("File size: %s, width: %d, height: %d, records: %u"), 
				PrintSize(pHdr->nBytes), 
				pHdr->rclBounds.right - pHdr->rclBounds.left,
				pHdr->rclBounds.bottom - pHdr->rclBounds.top,
				pHdr->nRecords);
			if (pHdr->nDescription)
			{
				LPTSTR lpszText = m_strText.GetBufferSetLength(pHdr->nDescription);
				::GetEnhMetaFileDescription(StgMedium.hEnhMetaFile, pHdr->nDescription, lpszText);
				m_strText.ReleaseBuffer();
			}
		}
		break;
	case CF_HDROP :							// List of files
		GetDropFileList(_T("files"));
		break;
	case CF_LOCALE :						// LCID for ANSI/OEM to Unicode conversions
		m_strDataType = _T("LCID");
		if (HasData(sizeof(LCID)))
		{
			SeparateContent();
			m_strContent.AppendFormat(_T("%#X"), *(static_cast<const LCID*>(m_lpData)));
		}
		break;
	case CF_DIB :							// device independant bitmap
	case CF_DIBV5 :
		m_strDataType = (CF_DIBV5 == m_FormatEtc.cfFormat) ? _T("BITMAPV5HEADER") : _T("BITMAPINFOHEADER");
		if (HasData(sizeof(BITMAPINFOHEADER)))
			GetBitmapInfo(static_cast<const BITMAPINFOHEADER*>(m_lpData));
		break;
	}
}

// Get information for common registered clipboard formats
void CDataInfo::GetCustomFormatInfo(COleDataObject* pDataObj, STGMEDIUM& StgMedium)
{
	// Microsoft HTML format (UTF-8 encoded)
	// See "HTML Clipboard Format": http://msdn.microsoft.com/en-us/library/Aa767917.aspx
	if (!m_strFormat.CompareNoCase(_T("HTML Format")))
		m_nCP = CP_UTF8;

	// CSV format (ANSI encoded)
	else if (!m_strFormat.CompareNoCase(_T("CSV")))
		m_nCP = CP_ACP;
//		m_nCP = m_pDropTarget->GetCodePage(pDataObj);

	// RTF formats (ASCII encoded), see richedit.h for definitions of format names
	else if (!m_strFormat.CompareNoCase(CF_RTF) ||
		!m_strFormat.CompareNoCase(CF_RTFNOOBJS) ||
		!m_strFormat.CompareNoCase(CF_RETEXTOBJ) ||
		!m_strFormat.CompareNoCase(_T("RTF As Text")))
	{
		m_nCP = 20127;
	}

	// Link related Microsoft formats

	// Link lists of null terminated strings (ANSI encoded)
	else if (!m_strFormat.CompareNoCase(_T("Link")) ||
		!m_strFormat.CompareNoCase(_T("ObjectLink")))
	{
		GetStringList(CP_ACP, _T("items"));
	}
	// CF_OBJECTDESCRIPTOR, CF_LINKSRCDESCRIPTOR: OBJECTDESCRIPTOR structure.
	else if (!m_strFormat.CompareNoCase(_T("Object Descriptor")) ||
		!m_strFormat.CompareNoCase(_T("Link Source Descriptor")))
	{
		GetObjectDescriptor();
	}
	// CF_LINKSOURCE: IMoniker followed by document class ID stored in IStream.
	else if (!m_strFormat.CompareNoCase(_T("Link Source")))
	{
		GetLinkSource(StgMedium);
	}
	// CFSTR_HYPERLINK: IHlink stored in IStream or HGLOBAL.
	else if (!m_strFormat.CompareNoCase(CFSTR_HYPERLINK))
	{
		GetHyperLink(pDataObj, StgMedium);
	}
			
	// Shell Object formats, see shellobj.h for definitions of format names.
	// MSDN: Shell Clipboard Formats
	//  http://msdn.microsoft.com/en-us/library/windows/desktop/bb776902%28v=vs.85%29.aspx
	// MSDN: Shell Data Object
	//  http://msdn.microsoft.com/en-us/library/windows/desktop/bb776903%28v=vs.85%29.aspx

	// CFSTR_FILENAME: Single file name or first name of list (ANSI encoded, short 8.3 name)
	// CFSTR_INETURLA: URL (ANSI encoded)
	else if (!m_strFormat.CompareNoCase(CFSTR_FILENAMEA) ||
		!m_strFormat.CompareNoCase(CFSTR_INETURLA))
		m_nCP = CP_ACP;

	// CFSTR_FILENAMEW: Single file name or first name of list (Unicode)
	// CFSTR_INETURLW: URL (Unicode)
	// CFSTR_MOUNTEDVOLUME: Path to mounted volume (Unicode)
	// CFSTR_INVOKECOMMAND_DROPPARAM: 
	else if (!m_strFormat.CompareNoCase(CFSTR_FILENAMEW) ||
		!m_strFormat.CompareNoCase(CFSTR_INETURLW) ||
		!m_strFormat.CompareNoCase(CFSTR_MOUNTEDVOLUME) ||
		!m_strFormat.CompareNoCase(CFSTR_INVOKECOMMAND_DROPPARAM))
		m_nCP = 1200;

	// CFSTR_FILEDESCRIPTOR: Virtual files, contents in CFSTR_FILECONTENTS with lindex
	else if (!m_strFormat.CompareNoCase(CFSTR_FILEDESCRIPTORA))
	{
		m_nCP = CP_ACP;
		m_strDataType = _T("FILEGROUPDESCRIPTORA");
		if (HasData(sizeof(FILEGROUPDESCRIPTORA)))
		{
			const FILEGROUPDESCRIPTORA *lpFGP = static_cast<const FILEGROUPDESCRIPTORA*>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("%u files"), lpFGP->cItems);
			if (HasAllData())
			{
				for (UINT i = 0; i < lpFGP->cItems; i++)
				{
					if (i)
						m_strText += _T(", ");
#ifdef _UNICODE
					m_strText += CString(lpFGP->fgd[i].cFileName);
#else
					m_strText += lpFGP->fgd[i].cFileName;
#endif
				}
			}
		}
	}
	else if (!m_strFormat.CompareNoCase(CFSTR_FILEDESCRIPTORW))
	{
		m_nCP = 1200;
		m_strDataType = _T("FILEGROUPDESCRIPTORW");
		if (HasData(sizeof(FILEGROUPDESCRIPTORW)))
		{
			const FILEGROUPDESCRIPTORW *lpFGP = static_cast<const FILEGROUPDESCRIPTORW*>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("%u files"), lpFGP->cItems);
			if (HasAllData())
			{
				for (UINT i = 0; i < lpFGP->cItems; i++)
				{
					if (i)
						m_strText += _T(", ");
#ifdef _UNICODE
					m_strText += lpFGP->fgd[i].cFileName;
#else
					m_strText += CString(lpFGP->fgd[i].cFileName);
#endif
				}
			}
		}
	}

	// CFSTR_FILECONTENTS: Virtual file content; may be HGLOBAL, IStream, and IStorage
	else if (!m_strFormat.CompareNoCase(CFSTR_FILECONTENTS))
	{
		GetFileContentsInfo(pDataObj);
	}

	// CFSTR_FILENAMEMAP: List of file names when renaming with HDROP
	else if (!m_strFormat.CompareNoCase(CFSTR_FILENAMEMAPA))
	{
		GetStringList(CP_ACP, _T("files"));
	}
	else if (!m_strFormat.CompareNoCase(CFSTR_FILENAMEMAPW))
	{
		GetStringList(1200, _T("files"));
	}

	// CFSTR_SHELLIDLIST: Locations of namespace objects
	else if (!m_strFormat.CompareNoCase(CFSTR_SHELLIDLIST) ||
		!m_strFormat.CompareNoCase(CFSTR_AUTOPLAY_SHELLIDLISTS))
	{
		m_strDataType = _T("CIDA");
		if (HasData(sizeof(_IDA)))
		{
			const _IDA* lpIda = static_cast<const _IDA*>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("%u PIDLs"), lpIda->cidl);
		}
	}

	// CFSTR_SHELLIDLISTOFFSET: Group and object positions for 
	//  CF_HDROP, CFSTR_SHELLIDLIST, and CFSTR_FILECONTENTS.
	else if (!m_strFormat.CompareNoCase(CFSTR_SHELLIDLISTOFFSET))
	{
		m_strDataType = _T("POINT[]");
		if (HasData(sizeof(POINT)))
		{
			const POINT *lpPt = static_cast<const POINT*>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("Left: %d top: %d"), lpPt->x, lpPt->y);
		}
	}

	// CFSTR_TARGETCLSID: Drop target CLSID
	// CFSTR_SHELLDROPHANDLER: CLSID of drop handler
	else if (!m_strFormat.CompareNoCase(CFSTR_TARGETCLSID) ||
		!m_strFormat.CompareNoCase(CFSTR_SHELLDROPHANDLER))
	{
		m_strDataType = _T("CLSID");
		if (HasData(sizeof(CLSID)))
			AppendCLSID(*(static_cast<const CLSID*>(m_lpData)), NULL);
	}

	// CFSTR_NETRESOURCES
	else if (!m_strFormat.CompareNoCase(CFSTR_NETRESOURCES))
	{
		m_strDataType = _T("NRESARRAY");
		if (HasData(sizeof(NRESARRAY)))
		{
			const NRESARRAY *lpNRA = static_cast<const NRESARRAY*>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("%u network resources"), lpNRA->cItems);
		}
	}

	// CFSTR_PRINTERGROUP: DROPFILES data with list of printer names
	else if (!m_strFormat.CompareNoCase(CFSTR_PRINTERGROUP))
	{
		GetDropFileList(_T("printers"));
	}

	// CFSTR_FILE_ATTRIBUTES_ARRAY: 
	else if (!m_strFormat.CompareNoCase(CFSTR_FILE_ATTRIBUTES_ARRAY))
	{
		m_strDataType = _T("FILE_ATTRIBUTES_ARRAY");
		if (HasData(sizeof(FILE_ATTRIBUTES_ARRAY)))
		{
			const FILE_ATTRIBUTES_ARRAY *lpFAA = static_cast<const FILE_ATTRIBUTES_ARRAY*>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("%u items"), lpFAA->cItems);
		}
	}

	// CFSTR_DROPDESCRIPTION: Cursor image type and optional text description
	else if (!m_strFormat.CompareNoCase(CFSTR_DROPDESCRIPTION))
	{
		m_strDataType = _T("DROPDESCRIPTION");
		if (HasData(sizeof(DROPDESCRIPTION)))
		{
			const DROPDESCRIPTION *lpDS = static_cast<const DROPDESCRIPTION*>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("Image type %d, szMessage: \"%ls\", szInsert: \"%ls\""), 
				lpDS->type, lpDS->szMessage, lpDS->szInsert);
		}
	}

	// DragImageBits: Drag image; SHDRAGIMAGE struct followed by the image bits.
	else if (!m_strFormat.CompareNoCase(_T("DragImageBits")))
	{
		m_strDataType = _T("SHDRAGIMAGE");
		if (HasData(sizeof(SHDRAGIMAGE)))
		{
			const SHDRAGIMAGE *lpDI = static_cast<const SHDRAGIMAGE*>(m_lpData);
			SeparateContent();
			m_strContent.AppendFormat(_T("Width: %d, height: %d, transparent color: %#010X"), 
				lpDI->sizeDragImage.cx, lpDI->sizeDragImage.cy, lpDI->crColorKey);
		}
	}

	// MIME formats (e.g. used by Mozilla applications)
	// Note that case sensitive compare is used.
	// While clipboard format names are not case sensitive, MIME formats
	//  must be lowercase.
	else if (!_tcsncmp(m_strFormat.GetString(), _T("text/"), 5))
	{
		CheckForString();
		// Assume Unicode when not identified as ASCII, ANSI, or UTF-8.
		if ((UINT)-1 == m_nCP)
			m_nCP = 1200;
	}

	// Handle some known BOOL, DWORD, and DROPEFFECT data types.
	else if (4 == m_nSize)
	{
		CString strExtra;
		DWORD dwData = HasData(sizeof(DWORD)) ? *(static_cast<const DWORD*>(m_lpData)) : 0;
		// DROPEEFECT values
		if (!m_strFormat.CompareNoCase(CFSTR_PREFERREDDROPEFFECT) ||
			!m_strFormat.CompareNoCase(CFSTR_LOGICALPERFORMEDDROPEFFECT) ||
			!m_strFormat.CompareNoCase(CFSTR_PASTESUCCEEDED) ||
			!m_strFormat.CompareNoCase(CFSTR_PERFORMEDDROPEFFECT))
		{
			m_strDataType = _T("DROPEEFECT");
			strExtra = GetDropEffectString(dwData);
		}
		// Boolean values
		// IsShowingLayered: True if target shows the drag image (calls IDropTarget handlers)
		// IsShowingText: True when the target window is actually showing drop descriptions
		// UsingDefaultDragImage: Set when drag image is created by the system
		// DisableDragText: ?, set when dragging from Explorer
		// IsComputingImage: ?, set when dragging from Explorer
		else if (!m_strFormat.CompareNoCase(CFSTR_INDRAGLOOP) ||
			!m_strFormat.CompareNoCase(_T("UsingDefaultDragImage")) ||
			!m_strFormat.CompareNoCase(_T("IsShowingLayered")) ||
			!m_strFormat.CompareNoCase(_T("IsShowingText")) ||
			!m_strFormat.CompareNoCase(_T("DisableDragText")) ||
			!m_strFormat.CompareNoCase(_T("IsComputingImage")))
		{
			m_strDataType = _T("BOOL");
			strExtra = dwData ? _T("TRUE") : _T("FALSE");
		}
		// Parameter value passed to IDragSourceHelper2::SetFlags()
		else if (!m_strFormat.CompareNoCase(_T("DragSourceHelperFlags")))
		{
			m_strDataType = _T("DWORD");
			if (DSH_ALLOWDROPDESCRIPTIONTEXT == dwData)
				strExtra = _T("DSH_ALLOWDROPDESCRIPTIONTEXT");
		}
		// DragWindow: HWND of drag image window (always a DWORD, even with 64-bit apps!)
		else if (!m_strFormat.CompareNoCase(_T("DragWindow")))
			m_strDataType = _T("DWORD (casted HWND)");
		// CFSTR_UNTRUSTEDDRAGDROP: URL action flag
		// ComputedDragImage: ?, set when dragging from Explorer
		else if (!m_strFormat.CompareNoCase(CFSTR_UNTRUSTEDDRAGDROP) ||
			!m_strFormat.CompareNoCase(_T("ComputedDragImage")))
		{
			m_strDataType = _T("DWORD");
		}
		// Format has been identified.
		if (HasData(sizeof(DWORD)) && !m_strDataType.IsEmpty())
		{
			SeparateContent();
			m_strContent.AppendFormat(_T("%#010X"), dwData);
			if (!strExtra.IsEmpty())
				m_strContent.AppendFormat(_T(" (%s)"), strExtra.GetString());
		}
	}
}

// Identify some common binary types
bool CDataInfo::CheckBinaryFormats()
{
	if (m_nBufSize >= 16)
	{
		const BYTE *lpData = GetByteData();
		LPCSTR lpszData = reinterpret_cast<LPCSTR>(lpData);
		// BMP bitmap file
		if (0 == memcmp(lpData, "BM", 2) && 0 == memcmp(lpData + 6, "\0\0\0", 4))
		{
			m_strDataType = _T("BITMAPFILEHEADER");
			if (HasData(sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER)))
				GetBitmapInfo(reinterpret_cast<const BITMAPINFOHEADER*>(lpData + sizeof(BITMAPFILEHEADER)));
		}
		// ZIP file
		else if (0 == memcmp(lpData, "PK\x03\x04", 4))
		{
			m_strDataType = _T("ZIP archive");
			if (HasData(sizeof(ZipHdr_t)))
			{
				const ZipHdr_t *pZipHdr = static_cast<const ZipHdr_t*>(m_lpData);
				UINT nCP = (pZipHdr->wBitFlag & 0x0800) ? CP_UTF8 : 437;
#ifdef _UNICODE
				LPCTSTR lpszFileName = CDragDropHelper::MultiByteToWideChar(pZipHdr->lpszFileName, pZipHdr->wFileNameLength, nCP);
#else
				LPCTSTR lpszFileName = CDragDropHelper::MultiByteToMultiByte(pZipHdr->lpszFileName, pZipHdr->wFileNameLength, nCP);
#endif
				m_strContent.AppendFormat(_T("First file name: %s"), lpszFileName);
				delete [] lpszFileName;
			}
		}
		// RIFF / RIFF-WAVE
		else if (0 == memcmp(lpData, "RIFF", 4))
		{
			m_strDataType = _T("RIFF");
			RiffHdr_t *RiffHdr = (RiffHdr_t*)lpData;
			m_strContent.Format(_T("File size: %u"), RiffHdr->chunkSize + 8);
			// RIFF-WAVE
			if (0 == memcmp(RiffHdr->type, "WAVE", sizeof(RiffHdr->type)) &&
				0 == memcmp(RiffHdr->fmt, "fmt ", sizeof(RiffHdr->fmt)))
			{
				m_strDataType += _T("-WAVE");
				m_strContent.AppendFormat(
					_T(", Format: %06#X, channels: %u, rate: %s Hz"),
					RiffHdr->wf.wFormatTag,
					RiffHdr->wf.nChannels,
					PrintSize(RiffHdr->wf.nSamplesPerSec));
			}
		}
		// TIFF
		else if (0 == strcmp(lpszData, "II*") || 0 == strcmp(lpszData, "MM*"))
			m_strDataType = _T("TIFF");
		// PNG
		else if (0 == memcmp(lpData, "\x89PNG\r\n\x1A\n", 8))
			m_strDataType = _T("PNG");
		// GIF
		else if (0 == memcmp(lpData, "GIF87a", 6) || 0 == memcmp(lpData, "GIF89a", 6))
			m_strDataType = _T("GIF");
		// JPEG / Exif: 
		//  0xFF 0xD8 Start of image marker
		//  0xFF 0xEx APPx marker (e.g. E0 = JFIF, E1 = Exif)
		//  WORD: segment length
		//  char[5]: Null terminated identifier
		else if (0 == memcmp(lpData, "\xFF\xD8\xFF", 3) && 
			lpData[3] >= 0xE0 && lpData[3] <= 0xEF &&
			4 == strlen(lpszData + 6))
		{
			m_strDataType = lpszData + 6;
			m_strDataType += _T(" (JPEG)");
		}
		// OLE compound file (e.g. Excel BIFF)
		else if (0 == memcmp(lpData, "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1", 8))
			m_strDataType = _T("OLE compound file");
	}
	return !m_strDataType.IsEmpty();
}

// Append binary data to m_strContent
void CDataInfo::GetBinaryContent()
{
	SeparateContent();
	m_strContent += _T("Data: ");
	size_t nBytes = min(m_nBufSize, EXCERPT_LEN / 3);
	const BYTE *lpData = GetByteData();
	for (size_t i = 0; i < nBytes; i++)
		m_strContent.AppendFormat(_T("%02X "), lpData[i]);
	if (nBytes < m_nSize)
		m_strContent += _T("...");
}

void CDataInfo::GetTextInfo()
{
	// Extract begin of text (if text not already assigned).
	if (m_nCP != (UINT)-1 && m_strContent.IsEmpty() && m_strText.IsEmpty() && m_lpData && m_nBufSize)
	{
		size_t nLen = 0;
		if (1200 == m_nCP)
		{
			LPCWSTR lpszText = static_cast<LPCWSTR>(m_lpData);
			nLen = CDragDropHelper::GetWideLength(lpszText, m_nBufSize);
			int nExcerpt = min(EXCERPT_LEN, static_cast<int>(nLen));
			m_strText = CString(lpszText, nExcerpt);
		}
		else
		{
			LPCSTR lpszText = static_cast<LPCSTR>(m_lpData);
			nLen = CDragDropHelper::GetMultiByteLength(lpszText, m_nBufSize);
			int nExcerpt = min(EXCERPT_LEN, static_cast<int>(nLen));
#ifdef _UNICODE
			LPWSTR lpszExcerpt = CDragDropHelper::MultiByteToWideChar(lpszText, nExcerpt, m_nCP);
#else
			LPSTR lpszExcerpt = CDragDropHelper::MultiByteToMultiByte(lpszText, nExcerpt, m_nCP);
#endif
			m_strText = lpszExcerpt;
			delete [] lpszExcerpt;
		}
		// If the full text is available, show the number of characters.
		if (m_nBufSize == m_nSize)
		{
			// Get the number of characters for UTF-8 strings.
			if (CP_UTF8 == m_nCP)
				nLen = ::MultiByteToWideChar(CP_UTF8, 0, static_cast<LPCSTR>(m_lpData), static_cast<int>(nLen), NULL, 0);
			SeparateContent();
			m_strContent.AppendFormat(_T("%s characters"), PrintSize(nLen));
		}
	}

	// Create text encoding string from code page number.
	// m_nCP is from CF_LOCALE or set here to CP_ACP, 1200, 20127, or CP_UTF8.
	// So there is no need to assign strings for code pages that are not used by Windows
	//  like other UTF code pages (7, 16BE, 32).
	if (m_strDataType.IsEmpty() && (UINT)-1 != m_nCP)
	{
		if (CP_ACP == m_nCP || CP_OEMCP == m_nCP)
			m_nCP = CDragDropHelper::GetCodePage(LOCALE_SYSTEM_DEFAULT, CP_OEMCP == m_nCP); 
		switch (m_nCP)
		{
		case CP_ACP : m_strDataType = _T("ANSI"); break;
		case CP_OEMCP : m_strDataType = _T("OEM/DOS"); break;
		case 1200 : m_strDataType = _T("Unicode (CP 1200)"); break;
		case 20127 : m_strDataType = _T("ASCII (CP 20127)"); break;
		case CP_UTF8 : m_strDataType = _T("UTF-8 (CP 65001)"); break;
		default :
			if (m_nCP < 1200)
				m_strDataType.Format(_T("OEM/DOS (CP %d))"), m_nCP);
			else if (m_nCP < 10000)
				m_strDataType.Format(_T("ANSI (CP %d)"), m_nCP);
			else
				m_strDataType.Format(_T("CP %u"), m_nCP);
		}
	}

	// Append quoted strText to strContent.
	if (!m_strText.IsEmpty())
	{
		SeparateContent();
		m_strContent += _T('"');
		if (m_strText.GetLength() > EXCERPT_LEN)
		{
			m_strContent.Append(m_strText.GetString(), EXCERPT_LEN);
			m_strContent += _T("...");
		}
		else
			m_strContent += m_strText;
		m_strContent += _T('"');
	}
}

// Check if data is some kind of text.
// Return value:
// 0 = Not text.
// 1 = Not null terminated (may be false detected when m_nBufSize < m_nSize).
// 2 = Null terminated.
int CDataInfo::CheckForString()
{
	int nRet = 0;
	if (m_lpData && m_nBufSize)
	{
		LPCSTR lpszText = static_cast<LPCSTR>(m_lpData);
		bool bAscii = true;
		size_t nLen = 0;
		while (nLen < m_nBufSize && lpszText[nLen])
		{
			if (lpszText[nLen] & 0x80)
				bAscii = false;
			++nLen;
		}
		if ('\xFF' == lpszText[0] && '\xFE' == lpszText[1])
		{
			m_nCP = 1200;
			nRet = 2;
		}
		else if ('\xEF' == lpszText[0] && '\xBB' == lpszText[1] && '\xBF' == lpszText[2])
		{
			m_nCP = CP_UTF8;
			nRet = (nLen < m_nBufSize) ? 2 : 1;
		}
		else if (nLen >= m_nBufSize - 1)
		{
			if (bAscii)								// or UTF-8/ANSI with ASCII chars only
				m_nCP = 20127;
			else if (::MultiByteToWideChar(CP_UTF8, 0, lpszText, (int)nLen, NULL, 0) > 0)
				m_nCP = CP_UTF8;
			else
				m_nCP = CP_ACP;
			nRet = (nLen < m_nBufSize) ? 2 : 1;
		}
		else										// check for a wide string
		{
			LPCWSTR lpszWide = static_cast<LPCWSTR>(m_lpData);
			nLen = 0;
			size_t nSpaces = 0;
			size_t nMaxLen = m_nBufSize / sizeof(WCHAR);
			while (nLen < nMaxLen && lpszWide[nLen])
			{
				if (L' ' == lpszWide[nLen])
					++nSpaces;
				++nLen;
			}
			// NOTE: This fails with ANSI lists (multiple null terminated strings
			//  terminated by another null byte)
//			if (nLen == nMaxLen - 1)	// null terminated
//				m_nCP = 1200;
			// Check for a reasonable number of spaces if not null terminated.
//			else if (nLen == nMaxLen && nSpaces > nMaxLen / 15)
			if (nLen >= nMaxLen -1 && nSpaces > nMaxLen / 15)
			{
				m_nCP = 1200;
				nRet = (nLen < nMaxLen) ? 2 : 1;
			}
		}
	}
	return nRet;
}

void CDataInfo::GetBitmapInfo(const BITMAPINFOHEADER *lpInfo)
{
	SeparateContent();
	m_strContent.AppendFormat(_T("Image size: %s, width: %d, height: %d, color depth: %hu BPP"), 
		PrintSize(lpInfo->biSizeImage), 
		lpInfo->biWidth, lpInfo->biHeight, lpInfo->biBitCount);
}

// CF_OBJECTDESCRIPTOR, CF_LINKSRCDESCRIPTOR
void CDataInfo::GetObjectDescriptor()
{
	if (m_nSize)
		m_strDataType = _T("OBJECTDESCRIPTOR");
	if (HasData(sizeof(OBJECTDESCRIPTOR)))
	{
		const OBJECTDESCRIPTOR *pDescr = static_cast<const OBJECTDESCRIPTOR*>(m_lpData);
		if (HasAllData())
		{
			if (pDescr->dwFullUserTypeName)
			{
				SeparateContent();
				m_strContent.AppendFormat(_T("Type: %ls"),
					reinterpret_cast<LPCWSTR>(GetByteData() + pDescr->dwFullUserTypeName));
			}
			if (pDescr->dwSrcOfCopy)
			{
				SeparateContent();
				m_strContent.AppendFormat(_T("Source: %ls"),
					reinterpret_cast<LPCWSTR>(GetByteData() + pDescr->dwSrcOfCopy));
			}
		}
		AppendCLSID(pDescr->clsid, _T("CLSID"));
	}
}

// CF_LINKSOURCE data: IStream with IMoniker followed by document class ID.
void CDataInfo::GetLinkSource(STGMEDIUM& StgMedium)
{
	if (m_nSize)
	{
		m_strDataType = _T("IMoniker");
		if (TYMED_ISTREAM == StgMedium.tymed &&
			SUCCEEDED(StgMedium.pstm->Seek(m_LargeNull, STREAM_SEEK_SET, NULL)))
		{
			// Load object from stream
			LPMONIKER pMoniker = NULL;
			if (SUCCEEDED(::OleLoadFromStream(StgMedium.pstm, IID_IMoniker, (LPVOID*)&pMoniker)))
			{
				// Get the display name
				IBindCtx * pbc; 
				if (SUCCEEDED(::CreateBindCtx(NULL, &pbc)))
				{
					LPOLESTR lpOleStr = NULL;
					if (SUCCEEDED(pMoniker->GetDisplayName(pbc, NULL, &lpOleStr)))
						AppendOleStr(lpOleStr, NULL);
					pbc->Release();
				}
				// Read the class ID of the document that has been stored after the moniker.
				CLSID classID;
				if (SUCCEEDED(StgMedium.pstm->Read(&classID, sizeof(classID), NULL)))
					AppendCLSID(classID, _T("Document CLSID"));
				pMoniker->Release();
			}
		}
	}
}

// CF_HYPERLINK: IStream or HGLOBAL with IHlink.
void CDataInfo::GetHyperLink(COleDataObject* pDataObj, STGMEDIUM& StgMedium)
{
	LPHLINK pHlink = NULL;
	if (m_nSize)
	{
		m_strDataType = _T("IHlink");
		if (TYMED_ISTREAM == StgMedium.tymed)
		{
			// Rewind to begin of stream (actual position may be at end so that OleLoadFromStream() would fail)
			if (SUCCEEDED(StgMedium.pstm->Seek(m_LargeNull, STREAM_SEEK_SET, NULL)))
				::OleLoadFromStream(StgMedium.pstm, IID_IHlink, reinterpret_cast<LPVOID*>(&pHlink));
		}
		else if (TYMED_HGLOBAL == StgMedium.tymed)
		{
			// Creates IHlink if cfHyperlink, cfUniformResourceLocator, CF_HDROP, or cfFileName are provided.
			// Because we know that cfHyperlink is present, we can omit the call to HlinkQueryCreateFromData().
//			if (SUCCEEDED(::HlinkQueryCreateFromData(pDataObj->GetIDataObject(FALSE))))
			::HlinkCreateFromData(pDataObj->GetIDataObject(FALSE), 
				NULL, 0, NULL, IID_IHlink, 
				reinterpret_cast<LPVOID*>(&pHlink));
		}
	}
	if (pHlink)
	{
		LPOLESTR lpOleTarget = NULL;
		LPOLESTR lpOleLocation = NULL;
//		if (SUCCEEDED(pHlink->GetStringReference(HLINKGETREF_ABSOLUTE , &lpOleTarget, &lpOleLocation)))
		if (SUCCEEDED(pHlink->GetStringReference(HLINKGETREF_DEFAULT, &lpOleTarget, &lpOleLocation)))
		{
			AppendOleStr(lpOleTarget, _T("Target"));
			AppendOleStr(lpOleLocation, _T("Location"));
		}
		if (NULL == lpOleTarget && NULL == lpOleLocation)
		{
			if (SUCCEEDED(pHlink->GetFriendlyName(HLFNAMEF_TRYFULLTARGET, &lpOleTarget)))
				AppendOleStr(lpOleTarget, _T("URL"));
		}
		pHlink->Release();
	}
}

// Parse HDROP style file list (DROPFILES struct followed by null terminated names 
//  finally terminated by another null char).
// Returns the number of file names.
unsigned CDataInfo::GetDropFileList(LPCTSTR lpszType)
{
	unsigned nFiles = 0;
	m_strDataType = _T("DROPFILES");
	if (HasAllData())
	{
		const DROPFILES *lpDrop = static_cast<const DROPFILES*>(m_lpData);
		m_nCP = lpDrop->fWide ? 1200 : CP_ACP;
		nFiles = ParseStringList(GetByteData() + lpDrop->pFiles, lpszType);
	}
	return nFiles;
}

unsigned CDataInfo::GetStringList(UINT nCP, LPCTSTR lpszType)
{
	m_nCP = nCP;
	return HasAllData() ? ParseStringList(GetByteData(), lpszType) : 0;
}

// Parse list of null terminated strings followed by another terminating null character.
// Returns the number of strings.
unsigned CDataInfo::ParseStringList(const BYTE *p, LPCTSTR lpszType)
{
	ASSERT(p);
	ASSERT(lpszType);

	m_strText = _T("");
	unsigned nCount = 0;
	if (1200 == m_nCP)
	{
		LPCWSTR s = reinterpret_cast<LPCWSTR>(p);
		while (*s)
		{
			m_strText.AppendFormat(_T("%ls|"), s);
			s = 1 + wcschr(s, L'\0');
			++nCount;
		}
	}
	else
	{
		LPCSTR s = reinterpret_cast<LPCSTR>(p);
		while (*s)
		{
			m_strText.AppendFormat(_T("%hs|"), s);
			s = 1 + strchr(s, '\0');
			++nCount;
		}
	}
	m_strText += _T('|');
	SeparateContent();
	m_strContent.AppendFormat(_T("%u %s: "), nCount, lpszType); 
	if (m_strText.GetLength() > EXCERPT_LEN)
	{
		m_strContent += m_strText.Left(EXCERPT_LEN);
		m_strContent += _T(" ...");
	}
	else
		m_strContent += m_strText;
	m_strText = _T("");
	return nCount;
}

// Get info about file content from the descriptor
void CDataInfo::GetFileContentsInfo(COleDataObject *pDataObject)
{
	bool bUnicode = true;
	CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(CFSTR_FILEDESCRIPTORW);
	if (!pDataObject->IsDataAvailable(cfFormat))
	{
		bUnicode = false;
		cfFormat = CDragDropHelper::RegisterFormat(CFSTR_FILEDESCRIPTORA);
		if (!pDataObject->IsDataAvailable(cfFormat))
			cfFormat = 0;
	}
	// File descriptors should be stored in global memory.
	HGLOBAL hGlobal = cfFormat ? pDataObject->GetGlobalData(cfFormat) : NULL;
	if (hGlobal)
	{
		size_t nFileSize = 0;
		// If there is only one file, some applications did not set the index to zero!
		UINT nIndex = m_FormatEtc.lindex >= 0 ? m_FormatEtc.lindex : 0;
		CString strFileName;
		const CLSID *pClsid = NULL;
		if (bUnicode)
		{
			LPFILEGROUPDESCRIPTORW lpFGD = static_cast<LPFILEGROUPDESCRIPTORW>(::GlobalLock(hGlobal));
			if (nIndex < lpFGD->cItems)
			{
				const FILEDESCRIPTORW *lpDF = &lpFGD->fgd[nIndex];
				if (lpDF->dwFlags & FD_FILESIZE)
					nFileSize = lpDF->nFileSizeLow + ((unsigned __int64)lpDF->nFileSizeHigh << 32ULL);
				if (lpDF->dwFlags & FD_CLSID)
					pClsid = &lpDF->clsid;
				strFileName = lpDF->cFileName;
			}
		}
		else
		{
			LPFILEGROUPDESCRIPTORA lpFGD = static_cast<LPFILEGROUPDESCRIPTORA>(::GlobalLock(hGlobal));
			if (nIndex < lpFGD->cItems)
			{
				const FILEDESCRIPTORA *lpDF = &lpFGD->fgd[nIndex];
				if (lpDF->dwFlags & FD_FILESIZE)
					nFileSize = lpDF->nFileSizeLow + ((unsigned __int64)lpDF->nFileSizeHigh << 32ULL);
				if (lpDF->dwFlags & FD_CLSID)
					pClsid = &lpDF->clsid;
				strFileName = lpDF->cFileName;
			}
		}
		if (!strFileName.IsEmpty())
		{
			SeparateContent();
			m_strContent.AppendFormat(_T("File name: %s"), strFileName.GetString());
		}
		if (nFileSize)
		{
			SeparateContent();
			m_strContent.AppendFormat(_T("file size: %s"), PrintSize(nFileSize));
		}
		if (pClsid)
			AppendCLSID(*pClsid, _T("CLSID"));
		::GlobalUnlock(hGlobal);
		::GlobalFree(hGlobal);
	}
}

// Append CLSID to m_strContent
void CDataInfo::AppendCLSID(const CLSID& id, LPCTSTR lpszPrefix)
{
	if (id != CLSID_NULL)
	{
		LPOLESTR lpOleStr = NULL;
		if (SUCCEEDED(::StringFromCLSID(id, &lpOleStr)))
			AppendOleStr(lpOleStr, lpszPrefix);
	}
}

// Append OLESTR to m_strContent and free the string.
void CDataInfo::AppendOleStr(LPOLESTR lpOleStr, LPCTSTR lpszPrefix)
{
	if (lpOleStr)
	{
		SeparateContent();
		if (lpszPrefix)
			m_strContent.AppendFormat(_T("%s: "), lpszPrefix);
#ifdef _UNICODE
		m_strContent += lpOleStr;
#else
		m_strContent += CString(lpOleStr);
#endif
		::CoTaskMemFree(lpOleStr);
	}
}

// Print formatted size_t value with thousands separator to static buffer.
LPCTSTR CDataInfo::PrintSize(size_t nSize) const
{
	static TCHAR lpszBuf[32];

	size_t n2 = 0;
    size_t nScale = 1;
    while (nSize >= 1000) 
	{
        n2 += nScale * (nSize % 1000);
        nSize /= 1000;
        nScale *= 1000;
    }
	LPTSTR lpszOut = lpszBuf;
	int nBufSize = ARRAYSIZE(lpszBuf) ;
	int nPos = _stprintf_s(lpszOut, nBufSize, _T("%u"), nSize);
	//int nPos = _stprintf(lpszBuf, _T("%u"), nSize);
	while (nScale > 1 && nPos > 0 && nBufSize > nPos)
	{
		lpszOut += nPos;
		nBufSize -= nPos;
        nScale /= 1000;
        nSize = n2 / nScale;
        n2 %= nScale;
		nPos = _stprintf_s(lpszOut, nBufSize, _T("%s%03u"), m_lpszThSep, nSize);
		//nPos = _stprintf(lpszOut, _T("%s%03u"), m_lpszThSep, nSize);
	}
	return lpszBuf;
}

void CDataInfo::SeparateContent()
{
	if (!m_strContent.IsEmpty())
		m_strContent += _T(", ");
}

void CDataInfo::AppendString(CString& str, LPCTSTR lpszText, LPCTSTR lpszSep /*= _T(" | ")*/) const
{
	if (!str.IsEmpty())
		str += lpszSep;
	str += lpszText;
}

CString CDataInfo::GetDropEffectString(DROPEFFECT dwEffect) const
{
	if (DROPEFFECT_NONE == dwEffect)
		return _T("DROPEFFECT_NONE");
	CString str;
	if (dwEffect & DROPEFFECT_COPY)
		str = _T("DROPEFFECT_COPY");
	if (dwEffect & DROPEFFECT_MOVE)
		AppendString(str, _T("DROPEFFECT_MOVE"));
	if (dwEffect & DROPEFFECT_LINK)
		AppendString(str, _T("DROPEFFECT_LINK"));
	if (dwEffect & DROPEFFECT_SCROLL)
		AppendString(str, _T("DROPEFFECT_SCROLL"));
	return str;
}

// Get storage medium string.
CString CDataInfo::GetTymedString(DWORD dwTymed) const
{
	if (0 == dwTymed)
		return _T("NULL");
	CString str;
	if (dwTymed & TYMED_HGLOBAL)
		str = _T("HGLOBAL");
	if (dwTymed & TYMED_FILE)
		AppendString(str, _T("FILE"));
	if (dwTymed & TYMED_ISTREAM)
		AppendString(str, _T("ISTREAM"));
	if (dwTymed & TYMED_ISTORAGE)
		AppendString(str, _T("ISTORAGE"));
	if (dwTymed & TYMED_GDI)
		AppendString(str, _T("GDI"));
	if (dwTymed & TYMED_MFPICT)
		AppendString(str, _T("MFPICT"));
	if (dwTymed & TYMED_ENHMF)
		AppendString(str, _T("ENHMF"));
	return str;
}

LPCTSTR CDataInfo::GetAspectString() const
{
	LPCTSTR lpszAspect = _T("");
	switch (m_FormatEtc.dwAspect)
	{
	case DVASPECT_CONTENT : lpszAspect = _T("CONTENT"); break;
	case DVASPECT_THUMBNAIL : lpszAspect = _T("THUMBNAIL"); break;
	case DVASPECT_ICON : lpszAspect = _T("ICON"); break;
	case DVASPECT_DOCPRINT : lpszAspect = _T("DOCPRINT"); break;
	}
	return lpszAspect;
}

CString CDataInfo::GetDeviceString() const
{
	CString str = _T("All");
	if (m_FormatEtc.ptd)
	{
		// DVTARGETDEVICE strings are null terminated ANSI strings
		const BYTE *pBase = reinterpret_cast<const BYTE *>(m_FormatEtc.ptd);
		WORD wOffset = m_FormatEtc.ptd->tdDeviceNameOffset ? m_FormatEtc.ptd->tdDeviceNameOffset : m_FormatEtc.ptd->tdDriverNameOffset;
		if (wOffset)
			str = reinterpret_cast<LPCSTR>(pBase + wOffset);
		else if (m_FormatEtc.ptd->tdExtDevmodeOffset)
		{
			const DEVMODEA *pDevMode = reinterpret_cast<const DEVMODEA*>(pBase + m_FormatEtc.ptd->tdExtDevmodeOffset);
#ifdef _UNICODE
			str = CStringA(reinterpret_cast<LPCSTR>(pDevMode->dmDeviceName), sizeof(pDevMode->dmDeviceName));
#else
			str.SetString(reinterpret_cast<LPCSTR>(pDevMode->dmDeviceName), sizeof(pDevMode->dmDeviceName));
#endif
		}
		else
			str = _T("Unknown");
	}
	return str;
}

// Get name of pre-defined clipboard format.
LPCTSTR CDataInfo::GetFixedFormatName(CLIPFORMAT cfFormat) const
{
	LPCTSTR lpszName = _T("");
	switch (cfFormat)
	{
	case CF_TEXT : lpszName = _T("CF_TEXT"); break;
	case CF_BITMAP : lpszName = _T("CF_BITMAP"); break;
	case CF_METAFILEPICT : lpszName = _T("CF_METAFILEPICT"); break;
	case CF_SYLK : lpszName = _T("CF_SYLK"); break;
	case CF_DIF : lpszName = _T("CF_DIF"); break;
	case CF_TIFF : lpszName = _T("CF_TIFF"); break;
	case CF_OEMTEXT : lpszName = _T("CF_OEMTEXT"); break;
	case CF_DIB : lpszName = _T("CF_DIB"); break;
	case CF_PALETTE : lpszName = _T("CF_PALETTE"); break;
	case CF_PENDATA : lpszName = _T("CF_PENDATA"); break;
	case CF_RIFF : lpszName = _T("CF_RIFF"); break;
	case CF_WAVE : lpszName = _T("CF_WAVE"); break;
	case CF_UNICODETEXT : lpszName = _T("CF_UNICODETEXT"); break;
	case CF_ENHMETAFILE : lpszName = _T("CF_ENHMETAFILE"); break;
	case CF_HDROP : lpszName = _T("CF_HDROP"); break;
	case CF_LOCALE : lpszName = _T("CF_LOCALE"); break;
	case CF_DIBV5 : lpszName = _T("CF_DIBV5"); break;
	case CF_OWNERDISPLAY : lpszName = _T("CF_OWNERDISPLAY"); break;
	case CF_DSPTEXT : lpszName = _T("CF_DSPTEXT"); break;
	case CF_DSPBITMAP : lpszName = _T("CF_DSPBITMAP"); break;
	case CF_DSPMETAFILEPICT : lpszName = _T("CF_DSPMETAFILEPICT"); break;
	case CF_DSPENHMETAFILE : lpszName = _T("CF_DSPENHMETAFILE"); break;
	default :
		if (m_FormatEtc.cfFormat >= CF_PRIVATEFIRST && m_FormatEtc.cfFormat <= CF_PRIVATELAST)
			lpszName = _T("CF_PRIVATExxx");
		else if (m_FormatEtc.cfFormat >= CF_GDIOBJFIRST && m_FormatEtc.cfFormat <= CF_GDIOBJLAST)
			lpszName = _T("CF_GDIOBJxxx");
	}
	return lpszName;
}

// Get name of clipboard format.
bool CDataInfo::SetFormatName(CLIPFORMAT cfFormat)
{
	if (cfFormat >= 0xC000)							// registered names
	{
		TCHAR lpszName[128];
		if (::GetClipboardFormatName(cfFormat, lpszName, ARRAYSIZE(lpszName)) > 0)
			m_strFormat = lpszName;
		else
			m_strFormat = _T("");
	}
	else
		m_strFormat = GetFixedFormatName(cfFormat);
	return !m_strFormat.IsEmpty();
}

