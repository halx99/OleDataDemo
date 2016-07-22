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

// MyStatic.cpp : 
//  Simple picture control to demonstrate drag & drop.

#include "stdafx.h"
#include "OleDataDemo.h"
#include "OleDataSourceEx.h"
#include "MyStatic.h"
#include ".\mystatic.h"

// CMyStatic

IMPLEMENT_DYNAMIC(CMyStatic, CStatic)
CMyStatic::CMyStatic()
{
	m_bAllowDropDesc = false;						// Use drop description text (Vista and later)
	m_bDelayRendering = true;						// Use delay rendering when providing drag & drop data
	m_bDragging = false;							// Internal var; set when this is drag & drop source
	m_nDragImageType = DRAG_IMAGE_FROM_CAPT;		// Drag image type
	m_nDragImageScale = 0;							// Drag image scaling
	m_nDropType = DropNothing;						// Cached drop source type
}

CMyStatic::~CMyStatic()
{
}


BEGIN_MESSAGE_MAP(CMyStatic, CStatic)
	ON_COMMAND(ID_EDIT_PASTE, OnEditPaste)
	ON_COMMAND(ID_EDIT_COPY, OnEditCopy)
	ON_MESSAGE(WM_APP_DRAG_ENTER, OnDragEnter)
	ON_MESSAGE(WM_APP_DRAG_OVER, OnDragOver)
	ON_MESSAGE(WM_APP_DROP_EX, OnDropEx)
	ON_WM_LBUTTONDOWN()
END_MESSAGE_MAP()

void CMyStatic::Init(int nResID /*= 0*/)
{
	// Must be a picture control.
	ASSERT(SS_BITMAP == (GetStyle() & SS_TYPEMASK));
	// We need notifications to get the drag handlers be called.
	ASSERT(GetStyle() & SS_NOTIFY);

	// m_pDropTarget is CWnd member that is set to the address of the COleDropTarget class
	// when registering. If it is not NULL, this class has been already registered.
    if (NULL == m_pDropTarget)
    {
		// Use messages to pass drag events to this class.
		m_DropTarget.SetUseMsg();
		// Set drop description text.
		// Insert string is set by this class to the type of drop data. 
		m_DropTarget.SetDropDescriptionText(DROPIMAGE_MOVE, NULL, _T("Move %1"));
		m_DropTarget.SetDropDescriptionText(DROPIMAGE_COPY, NULL, _T("Copy %1"));
		VERIFY(m_DropTarget.Register(this)); 
	}
	if (nResID)
		LoadBitmap(nResID);
}

#if 0
// Register this window as drop target.
// OnCreate() is not called for template created controls.
// So it can be done here.
void CMyStatic::PreSubclassWindow()
{
	Init();
	CStatic::PreSubclassWindow();
}
#endif

// Check if file name has a supported image type extension.
bool CMyStatic::IsSupportedFileExtension(LPCTSTR lpszFileName) const
{
	bool bIsSupportedFileExtension = false;
	LPCTSTR lpszExt = ::PathFindExtension(lpszFileName);
	if (_T('.') == *lpszExt)
	{
#if 1
		++lpszExt;
		static LPCTSTR lpszExtensions[] = { 
			_T("bmp"), _T("dib"), _T("rle"), 
			_T("jpg"), _T("jpeg"), _T("jpe"), _T("jfif"), 
			_T("gif"), 
			_T("tif"), _T("tiff"), 
			_T("png"), 
			_T("ico") 
		};
		int i = 0;
		while (!bIsSupportedFileExtension && lpszExtensions[i])
		{
			if (0 == _tcsicmp(lpszExt, lpszExtensions[i]))
				bIsSupportedFileExtension = true;
			++i;
		}
#else
		// Check file name extension against those of supported GDI+ codecs.
		// Based on code form atlimage.h.
		// Note that this check does not include .ico files which are supported for loading.
		ULONG_PTR gdiplusToken;
		Gdiplus::GdiplusStartupInput input;
		if (Gdiplus::Ok == Gdiplus::GdiplusStartup(&gdiplusToken, &input, NULL))
		{
			UINT nEncoders;
			UINT nBytes;
			if (Gdiplus::Ok == Gdiplus::GetImageEncodersSize(&nEncoders, &nBytes))
			{
				LPBYTE pBuf = new BYTE[nBytes];
				Gdiplus::ImageCodecInfo* pEncoders = reinterpret_cast<Gdiplus::ImageCodecInfo*>(pBuf);
				if (Gdiplus::Ok == Gdiplus::GetImageEncoders(nEncoders, nBytes, pEncoders))
				{
					CStringW strExtW(lpszExt);
					for (UINT iCodec = 0; iCodec < nEncoders && !bIsSupportedFileExtension; iCodec++)
					{
						CStringW strExtensions(pEncoders[iCodec].FilenameExtension);
						int iStart = 0;
						do
						{
							CStringW strExtension = ::PathFindExtensionW(strExtensions.Tokenize(L";", iStart));
							if (iStart != -1 &&
								0 == strExtension.CompareNoCase(strExtW.GetString()))
							{
								bIsSupportedFileExtension = true;
							}
						} while (iStart != -1 && !bIsSupportedFileExtension);
					}
				}
				delete [] pBuf;
			}
			Gdiplus::GdiplusShutdown(gdiplusToken);
		}
#endif
	}
	return bIsSupportedFileExtension;
}

// Check if a single file name is provided that seems to be an image file.
bool CMyStatic::CanLoadFromFile(COleDataObject *pDataObject) const
{
	// Check if there is a HDROP object with a single file name.
	bool bCanLoadFromFile = (1 == m_DropTarget.GetFileNameCount(pDataObject));
	if (bCanLoadFromFile)
	{
		bCanLoadFromFile = false;
		CString str = m_DropTarget.GetSingleFileName(pDataObject);
		// Check if file exists
		if (!str.IsEmpty() &&
			INVALID_FILE_ATTRIBUTES != ::GetFileAttributes(str.GetString()))
		{
			// Check file extension.
			// This is a rather simple check.
			// A full check may get the file type by reading the first bytes from the file.
			bCanLoadFromFile = IsSupportedFileExtension(str.GetString());
		}
	}
	return bCanLoadFromFile;
}

// Check if drag source provides data that can be dropped.
// The type of data that would be used upon dropping is cached in m_nDropType.
LRESULT CMyStatic::OnDragEnter(WPARAM wParam, LPARAM lParam)
{
	m_nDropType = GetDropType(reinterpret_cast<COleDataObject*>(wParam), true);
	// Call OnDragOver() to get the drop effect.
	return OnDragOver(wParam, lParam);
}

LRESULT CMyStatic::OnDragOver(WPARAM /*wParam*/, LPARAM lParam)
{
	COleDropTargetEx::DragParameters *pParams = reinterpret_cast<COleDropTargetEx::DragParameters*>(lParam);
	DROPEFFECT dropEffect = DROPEFFECT_NONE;
	if (DropType::DropNothing != m_nDropType)
	{
		if (DropType::DropBitmap == m_nDropType)
			dropEffect = (pParams->dwKeyState & MK_CONTROL) ? DROPEFFECT_COPY : DROPEFFECT_MOVE;
		else // With file images and the drag image, only copying is supported here.
			dropEffect = DROPEFFECT_COPY;
		// Use drop description text to indicate what is being dropped.
		// The description is updated by COleDropTragetEx::OnDragOver().
		if (m_DropTarget.IsTextAllowed())
		{
			LPCWSTR lpszInsert = NULL;
			switch (m_nDropType)
			{
			case DropType::DropBitmap : lpszInsert = L"bitmap"; break;
			case DropType::DropImageFile : lpszInsert = L"image file content"; break;
			case DropType::DropDragImage : lpszInsert = L"drag image"; break;
			}
			m_DropTarget.SetDropInsertText(lpszInsert);
		}
	}
	return m_DropTarget.FilterDropEffect(dropEffect);
}

// Paste OLE drag & drop data.
// Dropping is handled by this function.
// So we don't need a OnDrop() function.
LRESULT CMyStatic::OnDropEx(WPARAM wParam, LPARAM lParam)
{
	// Update info list of main dialog.
	AfxGetMainWnd()->SendMessage(WM_USER_UPDATE_INFO, wParam, 0);

	DROPEFFECT nEffect = DROPEFFECT_NONE;
	COleDropTargetEx::DragParameters *pParams = reinterpret_cast<COleDropTargetEx::DragParameters*>(lParam);
	// pParams->dropEffect is the return value of OnDragOver() that has just been called.
	// Note that this may be DROPEFFECT_NONE.
	if (pParams->dropEffect & (DROPEFFECT_COPY | DROPEFFECT_MOVE))
	{
		if (PutOleData(reinterpret_cast<COleDataObject*>(wParam), m_nDropType))
			nEffect = pParams->dropEffect;
	}
	return nEffect;
}

// Check if data can be dropped or pasted and determine the type of data.
CMyStatic::DropType CMyStatic::GetDropType(COleDataObject *pDataObject, bool bDropDragImage /*= false*/) const
{
	DropType nDropType = DropType::DropNothing;
	if (!m_bDragging && IsWindowEnabled())
	{
		// Has some kind of bitmap
		if (pDataObject->IsDataAvailable(CF_BITMAP) ||
			pDataObject->IsDataAvailable(CF_DIB) ||
			pDataObject->IsDataAvailable(CF_DIBV5))
		{
			nDropType = DropType::DropBitmap;
		}
		// Has a single file name: Check if it is an image file.
		else if (CanLoadFromFile(pDataObject))
		{
			nDropType = DropType::DropImageFile;
		}
		// Check if there is a drag image
		else if (bDropDragImage &&
			pDataObject->IsDataAvailable(CDragDropHelper::RegisterFormat(_T("DragImageBits"))))
		{
			nDropType = DropType::DropDragImage;
		}
	}
	return nDropType;
}

// Get data and set bitmap when dropping or pasting.
bool CMyStatic::PutOleData(COleDataObject *pDataObject, DropType nType)
{
	bool bRet = false;
	switch (nType)
	{
	case DropType::DropBitmap :
		bRet = ReplaceBitmap(m_DropTarget.GetBitmap(pDataObject));
		break;
	case DropType::DropImageFile :
		bRet = LoadFromFile(m_DropTarget.GetSingleFileName(pDataObject).GetString());
		break;
	case DropType::DropDragImage :
		bRet = ReplaceBitmap(m_DropTarget.GetDragImageBitmap(pDataObject));
		break;
	}
	return bRet;
}

bool CMyStatic::CanPaste() const
{
	bool bRet = false;
	COleDataObject DataObj;
	if (DataObj.AttachClipboard())
		bRet = (DropType::DropNothing != GetDropType(&DataObj));
	return bRet;
}

void CMyStatic::OnEditPaste()
{
	COleDataObject DataObj;
	if (DataObj.AttachClipboard())
		PutOleData(&DataObj, GetDropType(&DataObj));
}

void CMyStatic::OnEditCopy()
{
	CacheOleData(GetBitmap(), true);
}

// Cache data for clipboard and drag & drop.
COleDataSourceEx *CMyStatic::CacheOleData(HBITMAP hBitmap, bool bClipboard)
{
	COleDataSourceEx *pDataSrc = NULL;
	if (hBitmap)
	{
		pDataSrc = new COleDataSourceEx(this);
		pDataSrc->CacheBitmap(hBitmap, bClipboard);
//		pDataSrc->CachePngAsStream(hBitmap);
		// Provide bitmap also as virtual file.
		if (!bClipboard)
			pDataSrc->CacheBitmapAsFile(hBitmap, _T("Image.bmp"));
		if (bClipboard)
		{
			pDataSrc->SetClipboard();
			pDataSrc = NULL;
		}
	}
	return pDataSrc;
}

// Set formats for delay rendering with and drag & drop.
COleDataSourceEx* CMyStatic::DelayRenderOleData(HBITMAP hBitmap)
{
	COleDataSourceEx *pDataSrc = NULL;
	if (hBitmap)
	{
		pDataSrc = new COleDataSourceEx(this, OnRenderData);

		// Passing FORMATETC is only necessary when allowing other tymeds than TYMED_HGLOBAL.
		FORMATETC fetc = { 0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL | TYMED_ISTREAM };
		pDataSrc->DelayRenderData(CF_DIBV5, &fetc);
		pDataSrc->DelayRenderData(CF_DIB, &fetc);

		// Must pass FORMATETC to set TYMED_GDI (defaults to TYMED_HGLOBAL when not passing this struct).
		fetc.tymed = TYMED_GDI;
		pDataSrc->DelayRenderData(CF_BITMAP, &fetc);

		// Converted image types (only IStream supported by COleDataSourceEx)
		fetc.tymed = TYMED_ISTREAM;
		pDataSrc->DelayRenderData(CF_TIFF, &fetc);
		pDataSrc->DelayRenderData(CDragDropHelper::RegisterFormat(_T("image/png")), &fetc);
		pDataSrc->DelayRenderData(CDragDropHelper::RegisterFormat(_T("image/gif")), &fetc);
		pDataSrc->DelayRenderData(CDragDropHelper::RegisterFormat(_T("image/jpeg")), &fetc);

		CLIPFORMAT cfContent = CDragDropHelper::RegisterFormat(CFSTR_FILECONTENTS);
		if (cfContent)
		{
			VirtualFileData VirtualData;			// Cache file descriptor using global data
			VirtualData.SetFileName(_T("Image.bmp"));
			if (pDataSrc->CacheFileDescriptor(1, &VirtualData))
			{
//				fetc.cfFormat = cfContent;			// Set by DelayRenderData()
				fetc.lindex = 0;					// Should pass FORMATETC to set the index
				fetc.tymed = TYMED_HGLOBAL | TYMED_ISTREAM;
				pDataSrc->DelayRenderData(cfContent, &fetc);
			}
		}
	}
	return pDataSrc;
}

// Static callback function to render data for drag & drop operations (when not using cached data)
BOOL CMyStatic::OnRenderData(CWnd *pWnd, const COleDataSourceEx *pDataSrc, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
	ASSERT_VALID(pWnd);
	ASSERT_VALID(pDataSrc);

	CMyStatic *pThis = static_cast<CMyStatic*>(pWnd);
	HBITMAP hBitmap = pThis->GetBitmap();
	// Render bitmap as CF_BITMAP, CF_DIB, CF_DIBV5, BMP image file content, or
	//  converted to PNG, TIFF, GIF, JPEG using IStream.
	return hBitmap ? pDataSrc->RenderBitmap(hBitmap, lpFormatEtc, lpStgMedium) : FALSE;
}

void CMyStatic::OnLButtonDown(UINT nFlags, CPoint point)
{
	bool bHandled = false;
	HBITMAP hBitmap = GetBitmap();
	COleDataSourceEx *pDataSrc = NULL;
	if (m_bDelayRendering)
		pDataSrc = DelayRenderOleData(hBitmap);
	else
		pDataSrc = CacheOleData(hBitmap, false);
	if (pDataSrc)
	{
		if (m_bAllowDropDesc)						// Allow targets to show user defined text.
			pDataSrc->AllowDropDescriptionText();	// Must be called before setting the drag image.

		CPoint pt(-1, 30000);						// Cursor is centered below the image
		switch (m_nDragImageType & DRAG_IMAGE_FROM_MASK)
		{
		case DRAG_IMAGE_FROM_CAPT :
			// Get drag image from content with optional scaling
			pDataSrc->InitDragImage(hBitmap, m_nDragImageScale, &pt);
			break;
		case DRAG_IMAGE_FROM_RES :
			// Use bitmap from resources as drag image
			pDataSrc->InitDragImage(IDB_BITMAP_DRAGME, &pt, CLR_INVALID);
			break;
		case DRAG_IMAGE_FROM_HWND :
			// This window handles DI_GETDRAGIMAGE messages.
//			pDataSrc->SetDragImageWindow(m_hWnd, &point);
//			break;
		case DRAG_IMAGE_FROM_SYSTEM :
			// Let Windows create the drag image.
			// With file lists and provided "Shell IDList Array" data, file type images are shown.
			// Otherwise, a generic image is used.
			pDataSrc->SetDragImageWindow(NULL, &point);
			break;
		case DRAG_IMAGE_FROM_EXT :
			// Get drag image from file extension using the file type icon
			pDataSrc->InitDragImage(_T(".bmp"), m_nDragImageScale, &pt);
			break;
//		case DRAG_IMAGE_FROM_TEXT :
//		case DRAG_IMAGE_FROM_BITMAP :
		default :
			// Get drag image from content without scaling
			pDataSrc->InitDragImage(hBitmap, 100, &pt);
		}
		m_bDragging = true;							// to know when dragging over our own window
		pDataSrc->DoDragDropEx(DROPEFFECT_COPY);
		pDataSrc->InternalRelease();
		m_bDragging = false;
		bHandled = true;
	}
	if (!bHandled)
		CStatic::OnLButtonDown(nFlags, point);
}

// Replace bitmap by another one.
bool CMyStatic::ReplaceBitmap(HBITMAP hBitmap)
{
	if (hBitmap)
	{
		HBITMAP hOld = SetBitmap(hBitmap);
		if (hOld)
			VERIFY(::DeleteObject(hOld));
	}
	return NULL != hBitmap;
}

// Load bitmap from resources.
bool CMyStatic::LoadBitmap(int nID)
{
	return ReplaceBitmap(::LoadBitmap(AfxGetResourceHandle(), MAKEINTRESOURCE(nID)));
}

// Load bitmap from image file.
bool CMyStatic::LoadFromFile(LPCTSTR lpszFileName)
{
	bool bRet = false;
	if (lpszFileName && *lpszFileName)
	{
		CImage Image;
		bRet = SUCCEEDED(Image.Load(lpszFileName));
		if (bRet)
			bRet = ReplaceBitmap(Image.Detach());
	}
	return bRet;
}

