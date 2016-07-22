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

// MyListCtrl.cpp: 
//  Report list control providing data for clipboard and OLE Drag&Drop.

#include "stdafx.h"
#include "resource.h"
#include "OleDataDemo.h"
#include "OleDataSourceEx.h"

#include ".\mylistctrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CGetTextHelper::CGetTextHelper()
{
	m_bAll = m_bHeader = m_bHidden = m_bHasCheckBoxes = false;
	m_nItemCount = m_nColumnCount = m_nSelectedCount = m_nUsedColumns = m_nTextSize = 0;
	m_pWidth = m_pOrder = m_pFormat = NULL;
};

CGetTextHelper::~CGetTextHelper()
{
	delete [] m_pFormat;
	delete [] m_pWidth;
	delete [] m_pOrder;
}

void CGetTextHelper::Clear()
{
	delete [] m_pFormat;
	delete [] m_pWidth;
	delete [] m_pOrder;
	m_bAll = m_bHeader = m_bHidden = m_bHasCheckBoxes = false;
	m_nItemCount = m_nColumnCount = m_nSelectedCount = m_nUsedColumns = m_nTextSize = 0;
	m_pWidth = m_pOrder = m_pFormat = NULL;
}

/////////////////////////////////////////////////////////////////////////////
// CMyListCtrl

IMPLEMENT_DYNAMIC(CMyListCtrl, CListCtrl)

CMyListCtrl::CMyListCtrl()
{
	m_bDragHighLight = true;						// Highlight cells below the drag cursor
	m_bDelayRendering = true;						// Use delay rendering when providing drag & drop data
	m_bCanDrop = false;								// Internal var; used with drop target
	m_bDragging = false;							// Internal var; set when this is drag & drop source
	m_bPadText = true;								// Pad plain text with spaces to get aligned output
	m_bAllowDropDesc = false;						// Use drop description text (Vista and later)
	m_nDragImageType = DRAG_IMAGE_FROM_TEXT;		// Drag image type
	m_nDragImageScale = 100;						// Drag image scaling factor
	m_cTextSep = _T(' ');							// Separation character for plain text data
	m_cCsvQuote = _T('"');							// Quote character for CSV data
	m_cCsvSep = _T(',');							// Separation character for CSV data
	m_bCsvQuoteSpace = false;						// Quote CSV data containing space or TAB
	m_bHtmlFont = true;								// Add font style to table tag
	m_nHtmlBorder = 1;								// Border width for HTML data
	m_nDragItem = m_nDragSubItem = -1;				// Item below drag cursor when highlighting cells
}

BEGIN_MESSAGE_MAP(CMyListCtrl, CListCtrl)
	ON_WM_CONTEXTMENU()
	ON_COMMAND(ID_EDIT_COPY, OnEditCopy)			// List context menu commands
#ifdef ID_EDIT_COPY_ALL // user defined
	ON_COMMAND(ID_EDIT_COPY_ALL, OnEditCopyAll)
#endif
	ON_COMMAND(ID_EDIT_SELECT_ALL, OnEditSelectAll)
#ifdef ID_EDIT_DESELECT // user defined
	ON_COMMAND(ID_EDIT_DESELECT, OnEditDeselect)
#endif
	// NOTE: Use ON_NOTIFY_REFLECT_EX here when message must be also handled by parent window.
	ON_NOTIFY_REFLECT(LVN_BEGINDRAG, OnLvnBegindrag)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnNMCustomdraw)
END_MESSAGE_MAP()


void CMyListCtrl::Init()
{
	// m_pDropTarget is CWnd member that is set to the address of the COleDropTarget class
	// when registering. If it is not NULL, this class has been already registered.
    if (NULL == m_pDropTarget)
    {
		// Additional styles
//		SetExtendedStyle(GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES | LVS_EX_HEADERDRAGDROP);
		SetExtendedStyle(GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT | LVS_EX_HEADERDRAGDROP);
		// Themed style list.
		if (::IsAppThemed() && (::GetThemeAppProperties() & STAP_ALLOW_CONTROLS))
			VERIFY(S_OK == ::SetWindowTheme(m_hWnd, L"Explorer", NULL));

		// Pass pointers to callback functions
		m_DropTarget.SetOnDragEnter(OnDragEnter);
		m_DropTarget.SetOnDragOver(OnDragOver);
		m_DropTarget.SetOnDropEx(OnDropEx);
		m_DropTarget.SetOnDragLeave(OnDragLeave);
		// Enable auto scrolling when dragging over scrollbars using callback function to perform the scrolling.
		m_DropTarget.SetScrollMode(AUTO_SCROLL_BARS | AUTO_SCROLL_DEFAULT, OnDoScroll);
//		m_DropTarget.SetDropDescriptionText(DROPIMAGE_MOVE, _T("Move"), _T("Move to %1"));
//		m_DropTarget.SetDropDescriptionText(DROPIMAGE_COPY, _T("Copy"), _T("Copy to %1"));
//		m_DropTarget.SetDropDescriptionText(DROPIMAGE_LINK, _T("Link"), _T("Set link into %1"));
		// Register window as drop target
		VERIFY(m_DropTarget.Register(this)); 
	}
}

#if 0
void CMyListCtrl::PreSubclassWindow()
{
	Init();
	CListCtrl::PreSubclassWindow();
}
#endif

// Static wrappers/handlers for drag support functions.

// Check if data can be pasted via OLE Drag&Drop.
// Text is pasted into cells and must not consist of multiple lines.
// The state is cached in m_bCanDrop to be used by OnDragOver().
DROPEFFECT CMyListCtrl::OnDragEnter(CWnd *pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point)
{
	ASSERT(pWnd->IsKindOf(RUNTIME_CLASS(CMyListCtrl)));
	CMyListCtrl *pThis = static_cast<CMyListCtrl*>(pWnd);
	unsigned nTextLines = 0;
	// Text length in characters for Unicode, ANSI, and OEM text and number of lines.
	// Zero if neither CF_TEXT, CF_OEMTEXT, nor CF_UNICODETEXT data available.
	size_t nLen = pThis->m_bDragging ? 0 : pThis->m_DropTarget.GetCharCount(pDataObject, &nTextLines);
	// Can drop if a single lined text.
	pThis->m_bCanDrop = (nLen > 0 && 1 == nTextLines);
	pThis->m_nDragItem = pThis->m_nDragSubItem = -1;
	return pThis->GetDropEffect(dwKeyState, point);
}

DROPEFFECT CMyListCtrl::OnDragOver(CWnd *pWnd, COleDataObject* /*pDataObject*/, DWORD dwKeyState, CPoint point)
{
	ASSERT(pWnd->IsKindOf(RUNTIME_CLASS(CMyListCtrl)));
	CMyListCtrl *pThis = static_cast<CMyListCtrl*>(pWnd);
	return pThis->GetDropEffect(dwKeyState, point);
}

DROPEFFECT CMyListCtrl::OnDropEx(CWnd *pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point)
{
	ASSERT(pWnd->IsKindOf(RUNTIME_CLASS(CMyListCtrl)));
	CMyListCtrl *pThis = static_cast<CMyListCtrl*>(pWnd);
	return pThis->OnDropEx(pDataObject, dropDefault, dropList, point);
}

void CMyListCtrl::OnDragLeave(CWnd *pWnd)
{
	ASSERT(pWnd->IsKindOf(RUNTIME_CLASS(CMyListCtrl)));
	CMyListCtrl *pThis = static_cast<CMyListCtrl*>(pWnd);
	pThis->RedrawOldDragItem();						// remove highlighting of cell below cursor
}

// Callback function to perform scrolling.
// Scrolling is detected by COleDropTargetEx and this is called only when content should be scrolled.
// nVert and nHorz are 0, 1, or -1 indicating the direction.
void CMyListCtrl::OnDoScroll(CWnd *pWnd, int nVert, int nHorz)
{
	ASSERT(pWnd->IsKindOf(RUNTIME_CLASS(CMyListCtrl)));
	CMyListCtrl *pThis = static_cast<CMyListCtrl*>(pWnd);
	pThis->RedrawOldDragItem();						// remove highlighting
	CRect rc(LVIR_BOUNDS, 0, 0, 0);					// get line height from rect of first item
	if (0 != pWnd->SendMessage(LVM_GETITEMRECT, 0, reinterpret_cast<LPARAM>(&rc)))
	{
		// Because LVM_SCROLL requires the scroll steps to be in pixels,
		//  the passed values must be multiplied with the line height and average char width.
		// But with list view style, the dx value specifies the number of columns to scroll.
		if (LVS_LIST != (pWnd->GetStyle() & LVS_TYPEMASK))
			nHorz *= 10;
		pWnd->SendMessage(LVM_SCROLL, nHorz, nVert * rc.Height());
	}
}

// Paste OLE drag & drop data.
// Dropping is handled by this function.
// So we don't need a OnDrop() function.
DROPEFFECT CMyListCtrl::OnDropEx(COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT /*dropList*/, CPoint point)
{
	// Update info list of main dialog.
	AfxGetMainWnd()->SendMessage(WM_USER_UPDATE_INFO, reinterpret_cast<WPARAM>(pDataObject), 0);

	RedrawOldDragItem();
	DROPEFFECT dwEffect = DROPEFFECT_NONE;
	if (dropDefault)
	{
		LVHITTESTINFO Info;
		Info.pt = point;
		Info.flags = LVHT_ONITEMLABEL;
		if (-1 != SubItemHitTest(&Info))			// get item and subitem below the mouse cursor
		{
			LPCTSTR lpszDropText = m_DropTarget.GetText(pDataObject);
			if (lpszDropText && SetItemText(Info.iItem, Info.iSubItem, lpszDropText))
				dwEffect = dropDefault;
			delete [] lpszDropText;
		}
	}
	return dwEffect;
}

// Get drop effect for OnDragEnter() and OnDragOver().
// Performed checks:
// - Check if over a used cell (not over header, unused cells, or scroll bar)
// - Determine the effect according to the key state
// - Optionally highlight the cell below the drag cursor
DROPEFFECT CMyListCtrl::GetDropEffect(DWORD dwKeyState, CPoint point)
{
	DROPEFFECT dropEffect = DROPEFFECT_NONE;
	if (m_bCanDrop)
	{
		LVHITTESTINFO Info;
		Info.pt = point;
		Info.flags = LVHT_ONITEMLABEL;
		if (-1 != SubItemHitTest(&Info))			// over a valid cell
		{
			bool bOnHeader = false;
			if (0 == Info.iItem && HasHeader())		// check if over the header
			{
				RECT rcHeader;
				GetHeaderCtrl()->GetClientRect(&rcHeader);
				if (point.y <= rcHeader.bottom)
					bOnHeader = true;
			}
			if (!bOnHeader)
				dropEffect = (dwKeyState & MK_CONTROL) ? DROPEFFECT_COPY : DROPEFFECT_MOVE;
		}
		if (m_bDragHighLight)
		{
			if (DROPEFFECT_NONE == dropEffect)
				RedrawOldDragItem();				// redraw cell from previous DragOver if necessary
			else if (m_nDragItem != Info.iItem ||	// over a new cell
				m_nDragSubItem != Info.iSubItem)
			{
				if (m_nDragItem != Info.iItem)		// previous cell was on another row
					RedrawOldDragItem();
				m_nDragItem = Info.iItem;
				m_nDragSubItem = Info.iSubItem;
				// This may be used to highlight entire rows.
//				SetItemState(m_nDragItem, LVIS_DROPHILITED, LVIS_DROPHILITED);
				RedrawItems(m_nDragItem, m_nDragItem);
			}
		}
	}
	return m_DropTarget.FilterDropEffect(dropEffect);
}

// Remove highlighting for previous cell.
void CMyListCtrl::RedrawOldDragItem()
{
	if (m_bDragHighLight && m_nDragItem >= 0)
	{
		// This may be used to highlight entire rows.
//		SetItemState(m_nDragItem, 0, LVIS_DROPHILITED);
		RedrawItems(m_nDragItem, m_nDragItem);
	}
	m_nDragItem = m_nDragSubItem = -1;
}

// Highlight cell when dragging over list.
// NOTE: 
//  Colors set here are ignored for hot and selected items with classic style.
//  With themed lists, other text colors are applied.
//  When also specifying the background color, such items may be shown with inverted
//   colors when using themes.
void CMyListCtrl::OnNMCustomdraw(NMHDR *pNMHDR, LRESULT *pResult)
{
	*pResult = CDRF_DODEFAULT;
	if (m_bDragHighLight)
	{
		LPNMLVCUSTOMDRAW lplvcd = reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR);
		switch (lplvcd->nmcd.dwDrawStage)
		{
		case CDDS_PREPAINT :						// first message at begin of each paint cycle
			*pResult = CDRF_NOTIFYITEMDRAW;			// notify of item specific paint operations
			break;
		case CDDS_ITEMPREPAINT :					// triggered by returning CDRF_NOTIFYITEMDRAW from CDDS_PREPAINT
			*pResult = CDRF_NOTIFYSUBITEMDRAW;		// handle each subitem seperately
			break;
		case CDDS_ITEMPREPAINT | CDDS_SUBITEM :		// triggered by returning CDRF_NOTIFYSUBITEMDRAW from CDDS_ITEMPREPAINT
			if (lplvcd->iSubItem == m_nDragSubItem &&
				lplvcd->nmcd.dwItemSpec == static_cast<DWORD>(m_nDragItem))
			{
				lplvcd->clrText = ::GetSysColor(COLOR_HOTLIGHT);
			}
			else
				lplvcd->clrText = GetTextColor();
			break;
		}
	}
}

// Cache data for clipboard and drag & drop.
COleDataSourceEx* CMyListCtrl::CacheOleData(bool bClipboard, bool bAll /*= false*/, bool bHeader /*= false*/, bool bHidden /*= false*/)
{
	CString strPlain, strCsv, strHtml;
	GetText(strPlain, strCsv, strHtml, bAll, bHeader, bHidden);
	if (strPlain.IsEmpty())
		return NULL;

	COleDataSourceEx *pDataSrc = new COleDataSourceEx(this);

	// Cache HTML, RTF, and CSV.
	pDataSrc->CacheHtml(strHtml.GetString());
	pDataSrc->CacheRtfFromText(strPlain.GetString(), this);
	pDataSrc->CacheCsv(strCsv.GetString());
	// Cache standard text formats and CF_LOCALE.
	pDataSrc->CacheText(strPlain.GetString(), bClipboard);
	// Cache selected lines as CF_BITMAP image.
	CBitmap Bitmap;
	if (GetBitmap(Bitmap, bAll, bHeader))
		pDataSrc->CacheDdb((HBITMAP)Bitmap.Detach(), NULL, false);
	// Cache virtual files when dragging.
	if (!bClipboard)
		CacheVirtualFiles(pDataSrc, strPlain, strCsv, strHtml);
	if (bClipboard)
	{
		pDataSrc->SetClipboard();
		pDataSrc = NULL;
	}
	return pDataSrc;
}

// Cache text data as virtual files.
void CMyListCtrl::CacheVirtualFiles(COleDataSourceEx *pDataSrc, const CString& strPlain, const CString& strCsv, const CString& strHtml)
{
	VirtualFileData FileData[3];
	FileData[0].SetFileName(_T("List_Content.txt"));
	FileData[1].SetFileName(_T("List_Content.csv"));
	FileData[2].SetFileName(_T("List_Content.html"));
	FileData[2].SetBOM("\xEF\xBB\xBF");			// prefix HTML with UTF-8 BOM
#ifdef _UNICODE
	CStringA strPlainA(strPlain);				// convert text and CSV to ANSI
	CStringA strCsvA(strCsv);
	FileData[0].SetString(strPlainA.GetString());
	FileData[1].SetString(strCsvA.GetString());
	LPSTR lpszHtml = CDragDropHelper::WideCharToMultiByte(strHtml.GetString(), -1, CP_UTF8);
#else
	FileData[0].SetString(strPlain.GetString());
	FileData[1].SetString(strCsv.GetString());
	LPSTR lpszHtml = CDragDropHelper::MultiByteToMultiByte(strHtml.GetString(), -1, CP_THREADACP, CP_UTF8);
#endif
	FileData[2].SetString(lpszHtml);
	pDataSrc->CacheVirtualFiles(3, FileData);
	delete [] lpszHtml;
}

// Set formats for delay rendering with drag & drop.
COleDataSourceEx* CMyListCtrl::DelayRenderOleData(bool bVirtualFiles)
{
	COleDataSourceEx *pDataSrc = new COleDataSourceEx(this, OnRenderData);

	// Passing FORMATETC is only necessary when allowing other tymeds than TYMED_HGLOBAL.
	FORMATETC fetc = { 0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL | TYMED_ISTREAM };

	pDataSrc->DelayRenderData(CDragDropHelper::RegisterFormat(CF_HTML_STR));
	pDataSrc->DelayRenderData(CDragDropHelper::RegisterFormat(CF_RTF_STR));
	pDataSrc->DelayRenderData(CDragDropHelper::RegisterFormat(CF_CSV_STR));

	// With clipboard operations, it is sufficient to pass only one text format.
	// When requesting other text formats, the system will provide a converted copy.
	// But this does not work with drag & drop operations.
	// So we must provide all necessary text formats:
	//  - CF_TEXT should be always provided (non Unicode applications usually did not support CF_UNICODETEXT)
	//  - CF_UNICODETEXT should be provided (at least when this is a Unicode application)
	//  - CF_OEMTEXT is optional
	pDataSrc->DelayRenderData(CF_UNICODETEXT, &fetc);
	pDataSrc->DelayRenderData(CF_TEXT, &fetc);
	pDataSrc->DelayRenderData(CF_OEMTEXT, &fetc);

	// Cache CF_LOCALE for text to Unicode conversions.
	pDataSrc->CacheLocale();

	// Provide selection as bitmap.
	// FORMATETC must be passed using TYMED_GDI. When passing NULL, DelayRenderData() uses TYMED_HGLOBAL.
	fetc.tymed = TYMED_GDI;
	pDataSrc->DelayRenderData(CF_BITMAP, &fetc);

	// TODO: This results in the last type only listed when enumerating formats!
	// VS 2003: COleDataSource::Lookup() checks if there is a matching entry in the cache.
	//  This checks the index but the index check is ored with a StgMedium.tymed == TYMED_NULL
	//   condition which is always true because the cached StgMedium member is cleared by 
	//   DelayRenderData() (NULL is used to indicate that the cache entry is for delay rendering
	//   and the data has not been cached yet). 
	// This is a MFC bug. There should be a special Lookup() function for delay rendering.
	// With cached data, the function is working as expected (because StgMedium contains data).
	// Pre VS 2003: The index check is not present.
	CLIPFORMAT cfContent = CDragDropHelper::RegisterFormat(CFSTR_FILECONTENTS);
	if (cfContent && bVirtualFiles)
	{
		VirtualFileData FileData[3];
		FileData[0].SetFileName(_T("List_Content.txt"));
		FileData[1].SetFileName(_T("List_Content.csv"));
		FileData[2].SetFileName(_T("List_Content.html"));
//		FileData[0].SetSize(GetPlainTextLength());
		if (pDataSrc->CacheFileDescriptor(3, FileData))
		{
			fetc.tymed = TYMED_HGLOBAL;
			for (fetc.lindex = 0; fetc.lindex < 3; fetc.lindex++)
			{
				pDataSrc->DelayRenderData(cfContent, &fetc);
			}
		}
	}
	return pDataSrc;
}

// Static callback function to render data for drag & drop operations (when not using cached data)
BOOL CMyListCtrl::OnRenderData(CWnd *pWnd, const COleDataSourceEx *pDataSrc, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
	ASSERT_VALID(pWnd);
	ASSERT_VALID(pDataSrc);

	bool bRet = false;
	CMyListCtrl *pThis = static_cast<CMyListCtrl*>(pWnd);

	if (CF_BITMAP == lpFormatEtc->cfFormat)
	{
		if (TYMED_NULL == lpStgMedium->tymed)		// can't reuse GDI objects
		{
			CBitmap Bitmap;
			bRet = pThis->GetBitmap(Bitmap, false, false);
			if (bRet)
			{
				lpStgMedium->tymed = TYMED_GDI;
				lpStgMedium->hBitmap = (HBITMAP)Bitmap.Detach();
				lpStgMedium->pUnkForRelease = NULL;
			}
		}
		return bRet;
	}

	bool bHtml = false;
	bool bCsv = false;
	bool bRtf = false;
	CLIPFORMAT cfFormat = lpFormatEtc->cfFormat;
	if (cfFormat >= 0xC000)
	{
		TCHAR lpszFormat[128] = _T("");
		::GetClipboardFormatName(cfFormat, lpszFormat, ARRAYSIZE(lpszFormat) );
		if (0 == _tcsicmp(lpszFormat, CFSTR_FILECONTENTS))
		{
			switch (lpFormatEtc->lindex)
			{
			case 0 :
				cfFormat = CF_TEXT;
				break;
			case 1 :
				bCsv = true;
				cfFormat = CDragDropHelper::RegisterFormat(CF_CSV_STR);
				break;
			case 2 :
				bHtml = true;
				cfFormat = CDragDropHelper::RegisterFormat(CF_HTML_STR);
				break;
			default :
				cfFormat = 0;
			}
		}
		else if (0 == _tcsicmp(lpszFormat, CF_HTML_STR))
			bHtml = true;
		else if (0 == _tcsicmp(lpszFormat, CF_RTF_STR))
			bRtf = true;
		else if (0 == _tcsicmp(lpszFormat, CF_CSV_STR))
			bCsv = true;
		else
			cfFormat = 0;							// unsupported registered clipboard format
	}
	CGetTextHelper Options;
	if (cfFormat && pThis->InitGetTextHelper(Options, false, false, false))
	{
		CString strData;
		if (bHtml)
			pThis->GetHtmlText(strData, Options);
		else if (bCsv)
			pThis->GetCsvText(strData, Options);
		else 
		{
			pThis->GetPlainText(strData, Options);
			if (bRtf)
			{
				CString strRtf;
				CTextToRTF Conv;
				if (NULL == Conv.TextToRtf(strRtf, strData.GetString(), pWnd))
					cfFormat = 0;
				else
					strData = strRtf;
			}
		}
		if (cfFormat)
			bRet = pDataSrc->RenderText(cfFormat, strData.GetString(), lpFormatEtc, lpStgMedium);
	}
	return bRet;
}

// Ole Drag&Drop of selected list rows.
// Provides data as Text, CSV and HTML.
void CMyListCtrl::OnLvnBegindrag(NMHDR *pNMHDR, LRESULT * /*pResult*/)
{
//	*pResult = 0;	// LVN_BEGINDRAG does not use the return value
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	if (LVIS_SELECTED == GetItemState(pNMLV->iItem, LVIS_SELECTED))
	{
		COleDataSourceEx *pDataSrc = NULL;
		if (m_bDelayRendering)
		{
			pDataSrc = DelayRenderOleData(true);
		}
		else
		{
			pDataSrc = CacheOleData(false);
		}
		if (NULL == pDataSrc)
			return;

		// Set drop description text. Must be called before setting the drag image.
		if (m_bAllowDropDesc)
		{
			pDataSrc->AllowDropDescriptionText();
			// Those image types that aren't set here will use the system default text.
//			pDataSrc->SetDropDescriptionText(DROPIMAGE_NONE, _T("Can't drop here"), NULL);
			pDataSrc->SetDropDescriptionText(DROPIMAGE_MOVE, _T("Move"), _T("Move to %1"));
			pDataSrc->SetDropDescriptionText(DROPIMAGE_COPY, _T("Copy"), _T("Copy to %1"));
			pDataSrc->SetDropDescriptionText(DROPIMAGE_LINK, _T("Link"), _T("Set link into %1"));
		}

		// Set drag image
		CPoint pt(-1, 30000);						// Cursor is centered below the image
		switch (m_nDragImageType & DRAG_IMAGE_FROM_MASK)
		{
		case DRAG_IMAGE_FROM_TEXT :
			// Create drag image from text string.
			{
				CString strPlain;
				if (GetPlainText(strPlain, false, false, false))
					pDataSrc->InitDragImage(this, strPlain.GetString(), NULL, 0 != (m_nDragImageType & DRAG_IMAGE_HITEXT));
				else
					pDataSrc->SetDragImageWindow(NULL, NULL);
			}
			break;
		case DRAG_IMAGE_FROM_RES :
			// Use bitmap from resources as drag image
			pDataSrc->InitDragImage(IDB_BITMAP_DRAGME, &pt, CLR_INVALID);
			break;
		case DRAG_IMAGE_FROM_HWND :
			// This window handles DI_GETDRAGIMAGE messages.
			// Supported by CListCtrl and CTreeViewCtrl.
			// This shows only the first column with report views!
			// Also, when dragging not from the leftmost cell, the image may be blended out.
			pDataSrc->SetDragImageWindow(m_hWnd, &pNMLV->ptAction);
			break;
		case DRAG_IMAGE_FROM_SYSTEM :
			// Let Windows create the drag image.
			// With file lists and provided "Shell IDList Array" data, file type images are shown.
			// Otherwise, a generic image is generated from the content.
//			pDataSrc->SetDragImageWindow(NULL, &pNMLV->ptAction);
			pDataSrc->SetDragImageWindow(NULL, NULL);
			break;
		case DRAG_IMAGE_FROM_EXT :
			// Get drag image from file extension using the file type icon.
			pDataSrc->InitDragImage(_T(".htm"), m_nDragImageScale, &pt);
			break;
		case DRAG_IMAGE_FROM_BITMAP :
			// Use special function to create drag image
			{
				CBitmap Bitmap;
				if (GetBitmap(Bitmap, false, false))
					pDataSrc->SetDragImage(Bitmap, &pt, GetBkColor());
//					pDataSrc->InitDragImage((HBITMAP)Bitmap, m_nDragImageScale, &pt, GetBkColor());
				else
					pDataSrc->SetDragImageWindow(NULL, NULL);
			}
			break;
		// case DRAG_IMAGE_FROM_CAPT :
		default :
			// Use the visible selected lines as drag image.
			// This may look uncommon if the selection is not a continous block of lines.
			// Then the image contains also unselected lines.
			int nFirst = -1;
			int nLast = -1;
			int nTop = GetTopIndex();
			int nBottom = nTop + GetCountPerPage();
			POSITION pos = GetFirstSelectedItemPosition();
			while (pos)
			{
				int nItem = GetNextSelectedItem(pos);
				if (nFirst < 0 && nItem >= nTop)
					nFirst = nLast = nItem;			// first visible and selected item
				else if (nFirst >= 0 && nItem < nBottom)
					nLast = nItem;					// last visible and selected item
			}
			CRect rect, rcItem;
			GetClientRect(&rect);					// list client rect
			VERIFY(GetItemRect(nFirst, &rcItem, LVIR_BOUNDS));
			rect.top = rcItem.top;					// top position from first item
			if (rcItem.right < rect.right)			// if used list size smaller than the window
				rect.right = rcItem.right;
			if (nFirst != nLast)
				VERIFY(GetItemRect(nLast, &rcItem, LVIR_BOUNDS));
			rect.bottom = rcItem.bottom;			// bottom position from last item
			pDataSrc->InitDragImage(this, &rect, m_nDragImageScale);
		}
		m_bDragging = true;							// to know when dragging over our own window
		CRect rcStart(0, 0, 0, 0);					// Pass a null rect to disable the start drag delay.
		pDataSrc->DoDragDropEx(DROPEFFECT_COPY, &rcStart);
		pDataSrc->InternalRelease();
		m_bDragging = false;
	}
}

// Copy current selection to the clipboard.
void CMyListCtrl::OnEditCopy()
{
	if (GetSelectedCount() > 0)
		CacheOleData(true);
}

// Copy all items to the clipboard including header and hidden columns.
void CMyListCtrl::OnEditCopyAll()
{
	if (GetItemCount() > 0)
		CacheOleData(true, true, true, true);
}

void CMyListCtrl::OnEditSelectAll()
{
	SetItemState(-1, LVIS_SELECTED, LVIS_SELECTED);
}

void CMyListCtrl::OnEditDeselect()
{
	SetItemState(-1, 0, LVIS_SELECTED);
}

void CMyListCtrl::OnContextMenu(CWnd* pWnd, CPoint point) 
{
#ifdef IDR_LIST_MENU								// List context menu resource ID
	if (point.x < 0 && point.y < 0)					// if opened by keyboard (not mouse)
	{
		int nItem = GetNextItem(-1, LVNI_FOCUSED);	// get focused item
		if (nItem < 0)								// no item has focus, so get first selected item
			nItem = GetNextItem(-1, LVNI_SELECTED);
		if (nItem >= 0 && GetSelectedCount())		// check if item is visible
		{											// if not, search for first selected and visible item
			int nTop = GetTopIndex();
			int nBottom = nTop + GetCountPerPage();
			if (nItem < nTop || nItem >= nBottom)
			{
				nItem = -1;
				POSITION pos = GetFirstSelectedItemPosition();
				while (pos && nItem < 0)				// use position of first visible selected item
				{
					int nSelItem = GetNextSelectedItem(pos);
					if (nSelItem >= nTop && nSelItem < nBottom)
						nItem = nSelItem;
				}
			}
		}
		if (nItem < 0 && !HasHeader())
			nItem = 0;
		RECT rect;
		if (nItem >= 0)								// position of selected or focused visible item
			GetItemRect(nItem, &rect, LVIR_LABEL);
		else										// position of header
			GetHeaderCtrl()->GetItemRect(0, &rect);
		point.SetPoint(0, rect.bottom);				// at left side of list below item or header
		ClientToScreen(&point);						// set absolute position for context menu
	}
	else
	{
		CPoint pt(point);							// Absolute point where clicked with mouse
		ScreenToClient(&pt);						// Get point relative to list.
		CRect rect;									// Check if clicked inside list area. 
		GetClientRect(&rect);						// Use client rect to exclude scroll bars.
		if (!rect.PtInRect(pt))						// Clicked outside the list area
		{											//  e.g. on a scroll bar
			CListCtrl::OnContextMenu(pWnd, point);	// Call default handler (e.g. context menu for scroll bars)
			return;
		}
	}
	CMenu Menu;
	Menu.LoadMenu(IDR_LIST_MENU);
	CMenu *pSubmenu = Menu.GetSubMenu(0);
	if (pSubmenu == NULL)
	{
		CListCtrl::OnContextMenu(pWnd, point);
		return;
	}

	if (0 == GetItemCount())
	{
		VERIFY(pSubmenu->EnableMenuItem(ID_EDIT_SELECT_ALL, MF_GRAYED) != -1);
#ifdef ID_EDIT_COPY_ALL
		VERIFY(pSubmenu->EnableMenuItem(ID_EDIT_COPY_ALL, MF_GRAYED) != -1);
#endif
	}
	if (0 == GetSelectedCount())
	{
		VERIFY(pSubmenu->EnableMenuItem(ID_EDIT_COPY, MF_GRAYED) != -1);
#ifdef ID_EDIT_DESELECT
		VERIFY(pSubmenu->EnableMenuItem(ID_EDIT_DESELECT, MF_GRAYED) != -1);
#endif
	}
	int nCmd = pSubmenu->TrackPopupMenuEx(			// Display the popup menu and return action.
		TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_VERPOSANIMATION | TPM_RETURNCMD | TPM_NONOTIFY,
		point.x, point.y, AfxGetMainWnd(), NULL);
	if (nCmd)
		SendMessage(WM_COMMAND, nCmd);
#else
	CListCtrl::OnContextMenu(pWnd, point);
#endif
}

// Create bitmap from list content.
// The purpose is to create a graphical representation of the list that must
//  not look exactly like lists on the screen.
// Parameters:
//  bAll:     Include all or selected lines.
//  bHeader:  Include header line if true.
//  nFlags:   Option to specify the look.
//   GET_BITMAP_HILIGHT  0x01  Selected items are drawn using system highlight colors (only if bAll is true)
//   GET_BITMAP_GRID     0x02  Draw grid lines
//   GET_BITMAP_BORDER   0x04  Draw border around list (1 pixel wide solid black line)
//   GET_BITMAP_HDRFRAME 0x08  Draw header items as buttons (normal text otherwise)
//
// Limitations:
// - Multi line text is not supported (neither for header nor for list cells)
// - Header images are not supported
// - Right aligned images are not supported
// - Themes are not supported (all drawing is performed using classic styles)
//   However, check boxes are drawn from the image list and will have the themed look therefore.
bool CMyListCtrl::GetBitmap(CBitmap& Bitmap, bool bAll /*= false*/, bool bHeader /*= false*/, unsigned nFlags /*= 0*/)
{
	ASSERT(NULL == Bitmap.m_hObject);				// bitmap handle attached to CBitmap

	const int nTextSpace = 4;						// horizontal spacing for cell text
	const int nImageSpace = 3;						// horizontal spacing for cell images

	DWORD dwView = GetView();
	int nItems = bAll ? GetItemCount() : GetSelectedCount();
	// We don't need all columns when not in report view and hidden columns are excluded
//	int nColumns = (bHidden || LV_VIEW_DETAILS == dwView) ? GetHeaderCtrl()->GetItemCount() : 1;
	int nColumns = (LV_VIEW_DETAILS == dwView) ? GetHeaderCtrl()->GetItemCount() : 1;
	if (0 == nItems || 0 == nColumns)
		return false;

	CRect rcItem;
	int nHeaderHeight = 0;							// get header height
	if (bHeader && LV_VIEW_DETAILS == dwView && HasHeader())
	{
		GetHeaderCtrl()->GetItemRect(0, rcItem);
		nHeaderHeight = rcItem.Height();
	}
	GetItemRect(0, &rcItem, LVIR_BOUNDS);			// get line height
	int nHeight = rcItem.Height();
//nFlags = 0xFF;

	CDC memDC;
	CClientDC dc(this);
	VERIFY(memDC.CreateCompatibleDC(&dc));
	CRect rcMem(0, 0, rcItem.Width(), nHeaderHeight + nItems * nHeight);
	bool bRet = (0 != Bitmap.CreateCompatibleBitmap(&dc, rcMem.Width(), rcMem.Height()));
	if (bRet)
	{
		CBitmap* pOldBitmap = memDC.SelectObject(&Bitmap);
		CFont *pOldFont = memDC.SelectObject(GetFont());
		memDC.SetTextColor(GetTextColor());
		memDC.SetBkColor(GetBkColor());
		memDC.FillSolidRect(&rcMem, dc.GetBkColor());

		DWORD dwExtendedStyle = GetExtendedStyle();
		bool bHasCheckBox = (0 != (dwExtendedStyle & LVS_EX_CHECKBOXES));
		bool bHasSubItemImages = (0 != (dwExtendedStyle & LVS_EX_SUBITEMIMAGES));
		int *pOrder = new int[nColumns];
		int *pWidth = new int[nColumns];
		int *pAlign = new int[nColumns];
		if (nColumns > 1)							// report view
		{
			GetColumnOrderArray(pOrder, nColumns);
			for (int i = 0; i < nColumns; i++)		// Get the width and alignment for each column
			{
				CRect rcColumn;
				GetHeaderCtrl()->GetItemRect(i, rcColumn);
				pWidth[i] = rcColumn.Width();
				// Column 0 is always left aligned in report view.
				pAlign[i] = i ? (GetColumnFormat(i) & LVCFMT_JUSTIFYMASK) : 0;
			}
		}
		else
		{
			pOrder[0] = 0;
			pWidth[0] = rcItem.Width();
			pAlign[0] = GetColumnFormat(0) & LVCFMT_JUSTIFYMASK;
		}
		// Get pointers to image lists and determine the sizes.
		CImageList *pCheckList = GetImageList(LVSIL_STATE);
		CImageList *pImageList = GetImageList(LV_VIEW_ICON == dwView || LV_VIEW_TILE == dwView ? LVSIL_NORMAL : LVSIL_SMALL);
		CSize szCheck(0, 0);
		CSize szImage(0, 0);
		if (pCheckList->GetSafeHandle() && bHasCheckBox)
		{
			int nCX, nCY;
			if (::ImageList_GetIconSize(pCheckList->GetSafeHandle(), &nCX, &nCY))
				szCheck.SetSize(nCX, nCY);
		}
		if (pImageList->GetSafeHandle())
		{
			int nCX, nCY;
			if (::ImageList_GetIconSize(pImageList->GetSafeHandle(), &nCX, &nCY))
				szImage.SetSize(nCX, nCY);
		}

		// Draw header (report view only)
		if (nHeaderHeight)
		{
			UINT nButtonStyle = DFCS_BUTTONPUSH;
			if (dwExtendedStyle & LVS_NOSORTHEADER)
				nButtonStyle |= DFCS_FLAT;
			CFont *pOldFont = memDC.SelectObject(GetHeaderCtrl()->GetFont());
			memDC.SetBkMode(TRANSPARENT);			// required when drawing buttons (nFlags & GET_BITMAP_HDRFRAME)
			CRect rcBtn(0, 0, 0, nHeaderHeight);
			for (int i = 0; i < nColumns; i++)
			{
				int nCol = pOrder[i];
				rcBtn.right = rcBtn.left + pWidth[nCol];
				if (nFlags & GET_BITMAP_HDRFRAME)
					memDC.DrawFrameControl(rcBtn, DFC_BUTTON, nButtonStyle);
				DrawBitmapText(memDC, GetColumnTitle(nCol), rcBtn, pAlign[nCol]);
				rcBtn.left = rcBtn.right;
			}
			memDC.SelectObject(pOldFont);
		}

		POSITION pos = GetFirstSelectedItemPosition();
		// When using the optional highlight colors, drawing mode must be set to TRANSPARENT.
		memDC.SetBkMode((bAll && (nFlags & GET_BITMAP_HILIGHT) && pos) ? TRANSPARENT : OPAQUE);
		rcItem.top = nHeaderHeight;
		for (int i = 0; i < nItems; i++)
		{
			rcItem.left = 0;
			rcItem.bottom = rcItem.top + nHeight;
			int nItem = bAll ? i : GetNextSelectedItem(pos);
			UINT nItemState = GetItemState(nItem, (UINT)-1);

			// Optional use highlight colors for selected rows (classic style).
			if (bAll && (nFlags & GET_BITMAP_HILIGHT))
			{
				if (nItemState & LVIS_SELECTED)
				{
					CRect rcRow(rcItem);
					// The background of check boxes including spacing is not set to COLOR_HIGHLIGHT.
					if (pCheckList && bHasCheckBox)
					{
						// When column 0 hasn't been reordered, just adjust the left position of the rect.
						// Otherwise, the left position of column 0 must be determined and used to draw
						//  two solid rects.
						if (pOrder[0])
						{
							CRect rcCol0;
							GetHeaderCtrl()->GetItemRect(0, rcCol0);
							rcRow.right = rcCol0.left;
							memDC.FillSolidRect(&rcRow, ::GetSysColor(COLOR_HIGHLIGHT));
							rcRow.left = rcCol0.left;
						}
						rcRow.left += min(nImageSpace + szCheck.cx + nTextSpace / 2, pWidth[0]);
					}
					rcRow.right = rcMem.right;
					memDC.FillSolidRect(&rcRow, ::GetSysColor(COLOR_HIGHLIGHT));
					memDC.SetTextColor(::GetSysColor(COLOR_HIGHLIGHTTEXT));
//					memDC.SetBkColor(::GetSysColor(COLOR_HIGHLIGHT));
				}
				else
				{
					memDC.SetTextColor(GetTextColor());
//					memDC.SetBkColor(GetBkColor());
				}
			}
			for (int j = 0; j < nColumns; j++)
			{
				int nCol = pOrder[j];
				rcItem.right = rcItem.left + pWidth[nCol];
				// Draw check box in column 0
				if (pCheckList && 0 == nCol && bHasCheckBox)
				{
					// Check box image states: 0 = hidden, 1 = unchecked, 2 = checked
					// Check box image indexes: 0 = unchecked, 1 = checked
					DrawBitmapImage(memDC, pCheckList, 
						((nItemState & LVIS_STATEIMAGEMASK) >> 12) - 1,
						rcItem, szCheck);
				}
				// Draw image
				// TODO: Test this with sub-item images
				if (pImageList && (0 == nCol || bHasSubItemImages))
				{
					LVITEM lvi;
					lvi.mask = LVIF_IMAGE;
					lvi.iItem = nItem;
					lvi.iSubItem = nCol;
					if (GetItem(&lvi))
						DrawBitmapImage(memDC, pImageList, lvi.iImage, rcItem, szImage);
				}
				// Draw the cell text
				DrawBitmapText(memDC, GetItemText(nItem, nCol), rcItem, pAlign[nCol]);
				rcItem.left = rcItem.right;			// left position for next column
			}
			rcItem.top = rcItem.bottom;				// top position for next row
		}
		// Draw grid lines.
		if ((nFlags & GET_BITMAP_GRID) && (dwExtendedStyle & LVS_EX_GRIDLINES))
		{
			CPen Pen(PS_SOLID, 1, memDC.GetNearestColor(0xF3F3F3));
			CPen *pOldPen = memDC.SelectObject(&Pen);
			for (int nY = nHeaderHeight ? nHeaderHeight : nHeight; nY < rcMem.bottom; nY += nHeight)
			{
				memDC.MoveTo(0, nY);
				memDC.LineTo(rcMem.right, nY);
			}
			for (int i = 1, nX = pWidth[pOrder[0]]; i < nColumns; i++)
			{
				memDC.MoveTo(nX, 0);
				memDC.LineTo(nX, rcMem.bottom);
				nX += pWidth[pOrder[i]];
			}
			VERIFY(&Pen == memDC.SelectObject(pOldPen));
		}
		// Draw border using the default 1 pixel wide solid black pen
		if (nFlags & GET_BITMAP_BORDER)
		{
			memDC.MoveTo(0, 0);
			memDC.LineTo(rcMem.right - 1, 0);
			memDC.LineTo(rcMem.right - 1, rcMem.bottom -1);
			memDC.LineTo(0, rcMem.bottom - 1);
			memDC.LineTo(0, 0);
		}
		delete [] pAlign;
		delete [] pWidth;
		delete [] pOrder;
		VERIFY(GetFont() == memDC.SelectObject(pOldFont));
		VERIFY(&Bitmap == memDC.SelectObject(pOldBitmap));
	}
	return bRet;
}

// Helper function for GetBitmap().
// Draw text for a single cell.
void CMyListCtrl::DrawBitmapText(CDC& memDC, const CString& str, const CRect& rcCell, int nAlign) const
{
	const int nTextSpace = 4;						// horizontal spacing for cell text
	CRect rc(rcCell);
	rc.DeflateRect(nTextSpace, 0);					// horizontal text spacing, text is drawn vertical centered
	if (rc.left < rc.right)							// if text can be shown
	{												//  (available space may be occupied by image or column is very small)
		CSize sz = memDC.GetTextExtent(str.GetString());
		UINT nFormat = DT_NOPREFIX | DT_WORD_ELLIPSIS | DT_SINGLELINE | DT_VCENTER;
		// Use left alignment if text does not fit in rectangle.
		if (sz.cx <= rc.Width())					// text will be not clipped
		{
			nFormat |= DT_NOCLIP;
			if (LVCFMT_RIGHT == nAlign)
				nFormat |= DT_RIGHT;
			else if (LVCFMT_CENTER == nAlign)
				nFormat |= DT_CENTER;
		}
		memDC.DrawText(str, &rc, nFormat);
	}
}

// Helper function for GetBitmap().
// When the image list is not empty, the left position of the passed rect
//  is adjusted by the width of the image and the image spacing to be used as start
//  position for following text.
void CMyListCtrl::DrawBitmapImage(CDC& memDC, CImageList *pImageList, int nImage, CRect& rc, const CSize& szImage) const
{
	const int nImageSpace = 3;						// horizontal spacing for cell images
	int nImageCount = pImageList ? pImageList->GetImageCount() : 0;
	if (nImageCount)
	{
		CSize sz;
		sz.cx = min(rc.Width() - nImageSpace, szImage.cx);
		if (sz.cx > 0)								// image can be (partially) drawn
		{
			if (nImage >= 0 && nImage < nImageCount)
			{
				sz.cy = min(rc.Height(), szImage.cy);
				POINT pt;
				pt.x = rc.left + nImageSpace;
				pt.y = rc.top + (rc.Height() - sz.cy + 1) / 2;
//				pImageList->DrawEx(&memDC, nImage, pt, sz, CLR_DEFAULT, CLR_DEFAULT, ILD_IMAGE);
				pImageList->DrawEx(&memDC, nImage, pt, sz, CLR_DEFAULT, CLR_DEFAULT, ILD_NORMAL);
			}
			rc.left += sz.cx + nImageSpace;			// position for text right of image (without text spacing)
		}
		else
			rc.left = rc.right;						// no space for image and text
	}
}

// Collect information used for copy and Drag & Drop operations.
int CMyListCtrl::InitGetTextHelper(CGetTextHelper& Options, bool bAll, bool bHeader, bool bHidden)
{
	// NOTE: GetView() requires a manifest specifying Comclt32.dll version 6.0!
	// With older versions, GetStyle() may be used to check for
	//  LVS_ICON, LVS_LIST, LVS_SMALLICON, and LVS_REPORT.
	DWORD dwView = GetView();

	Options.Clear();
	// Include hidden columns when requested and in report or list view
	Options.m_bHidden = bHidden && (LV_VIEW_DETAILS == dwView || LV_VIEW_LIST == dwView);
	Options.m_nItemCount = GetItemCount();
	Options.m_nSelectedCount = GetSelectedCount();
	// We don't need all columns when not in report view and hidden columns are excluded
	Options.m_nColumnCount = (Options.m_bHidden || LV_VIEW_DETAILS == dwView) ? GetHeaderCtrl()->GetItemCount() : 1;
	if (0 == Options.m_nItemCount || 0 == Options.m_nColumnCount || (0 == Options.m_nSelectedCount && !bAll))
		return 0;

	Options.m_bAll = bAll;
	// Include header when requested, style is report view, and header is present
	Options.m_bHeader = bHeader && LV_VIEW_DETAILS == dwView && HasHeader();
	Options.m_bHasCheckBoxes = (GetExtendedStyle() & LVS_EX_CHECKBOXES) != 0;
	Options.m_pOrder = new int[Options.m_nColumnCount];
	Options.m_pWidth = new int[Options.m_nColumnCount];
	Options.m_pFormat = new int[Options.m_nColumnCount];
	if (Options.m_nColumnCount > 1)
		GetColumnOrderArray(Options.m_pOrder, Options.m_nColumnCount);
	else
		Options.m_pOrder[0] = 0;

	// Get max. char width for each column and total number of chars.
	// Used to generate aligned text output with fixed size fonts and calculate the
	//  required buffer sizes.
	for (int j = 0; j < Options.m_nColumnCount; j++)
	{
		Options.m_pWidth[j] = 0;
		Options.m_pFormat[j] = GetColumnFormat(j);
		// GetColumnWidth() is for report and list styles only.
		if (Options.m_bHidden || 
			(LV_VIEW_DETAILS == dwView && GetColumnWidth(j)) || 
			(LV_VIEW_DETAILS != dwView && 0 == j))
		{
			++Options.m_nUsedColumns;
			if (Options.m_bHeader)					// get width of column header
			{
				Options.m_pWidth[j] = GetColumnTitle(j).GetLength();
				Options.m_nTextSize += Options.m_pWidth[j];
			}
			POSITION pos = GetFirstSelectedItemPosition();
			for (int i = 0; (i < Options.m_nItemCount) && (Options.m_bAll || pos); i++)
			{
				int nItem = Options.m_bAll ? i : GetNextSelectedItem(pos);
				int nWidth = GetItemText(nItem, j).GetLength();
				if (0 == j && Options.m_bHasCheckBoxes) 
					nWidth += 4;					// will prefix text with "[X] " or "[ ] "
				Options.m_nTextSize += nWidth;		// total char width
				if (nWidth > Options.m_pWidth[j])	// max. column char width
					Options.m_pWidth[j] = nWidth;
			}
		}
	}
	return Options.m_nTextSize;
}

// Get list content as text, CSV and HTML.
// bAll:    If true, copy all lines. If false, copy selected lines.
// bHeader: If true, first line is header line with column labels.
// bHidden: If true, copy also hidden columns (columns with width 0).
void CMyListCtrl::GetText(CString& strTxt, CString& strCsv, CString& strHtml, bool bAll /*= false*/, bool bHeader /*= false*/, bool bHidden /*= false*/)
{
	CGetTextHelper Options;
	if (InitGetTextHelper(Options, bAll, bHeader, bHidden))
	{
		GetPlainText(strTxt, Options);
		GetCsvText(strCsv, Options);
		GetHtmlText(strHtml, Options);
	}
}

// Get the content of all or selected lines into CString.
// Text is optional padded with spaces to get aligned text when using a non-proportional font.
int CMyListCtrl::GetPlainText(CString& str, bool bAll, bool bHeader, bool bHidden)
{
	CGetTextHelper Options;
	return InitGetTextHelper(Options, bAll, bHeader, bHidden) ? GetPlainText(str, Options) : 0;
}

int CMyListCtrl::GetPlainTextLength(const CGetTextHelper& Options) const
{
	ASSERT(Options.IsInitialized());

	int nSize = 0;
	if (m_bPadText)									// pad text with spaces
	{
		nSize = 2;									// line termination
		for (int i = 0; i < Options.m_nColumnCount; i++)
			nSize += Options.m_pWidth[i];
		nSize *= Options.GetTotalLines();			// multiply by number of lines
	}
	else
		nSize = Options.GetTextSizeLineFeed();		// plain text size including line feeds
	if (m_cTextSep)									// add size for separation characters
		nSize += Options.GetTotalColumnSeparatorCount();
	return nSize;
}

int CMyListCtrl::GetPlainText(CString& str, const CGetTextHelper& Options) const
{
	ASSERT(Options.IsInitialized());

	str.Empty();
	// Calculate the number of characters to be copied to the string.
	// With big lists, this avoids re-allocations.
	str.Preallocate(GetPlainTextLength(Options));

	if (Options.m_bHeader)							// header is also requested
		GetTextLine(str, -1, Options);
	if (Options.m_bAll)								// all lines
	{
		for (int i = 0; i < Options.m_nItemCount; i++)
			GetTextLine(str, i, Options);
	}
	else											// selected lines only
	{
		POSITION pos = GetFirstSelectedItemPosition();
		while (pos)
			GetTextLine(str, GetNextSelectedItem(pos), Options);
	}
	return str.GetLength();
}

// Append formatted list row (or header line) to strLine.
void CMyListCtrl::GetTextLine(CString& strLine, int nItem, const CGetTextHelper& Options) const
{
	bool bFirst = true;
	CString strItem;								// declare this outside the loop to avoid multiple allocations
	for (int i = 0; i < Options.m_nColumnCount; i++)
	{
		int j = Options.m_pOrder[i];				// Output according to the actual order on screen
		if (Options.m_pWidth[j])					// Hidden columns may be excluded
		{
			GetCellText(strItem, nItem, j, Options);	// Get text from header or sub-item
			if (!bFirst && m_cTextSep)
				strLine += m_cTextSep;				// separate columns
			int nPad = 0;
			if (m_bPadText)							// pad with spaces for aligned output
			{
				int nPadLeft = 0;
				int nFormat = Options.m_pFormat[j] & LVCFMT_JUSTIFYMASK;
				nPad = Options.m_pWidth[j] - strItem.GetLength();
				ASSERT(nPad >= 0);					// pWidth[j] has been initialized with invalid value
				if (nPad < 0)
					nPad = 0;
				if (LVCFMT_RIGHT == nFormat)
					nPadLeft = nPad;
				else if (LVCFMT_CENTER == nFormat)
					nPadLeft = nPad / 2;
				nPad -= nPadLeft;
				while (nPadLeft--)
					strLine += _T(' ');
			}
			strLine += strItem;
			while (nPad--)
				strLine += _T(' ');
			bFirst = false;
		}
	}
	strLine += _T("\r\n");							// separate rows by new line
}

// Get the content of all or selected lines into CSV formatted CString.
int CMyListCtrl::GetCsvText(CString& str, bool bAll, bool bHeader, bool bHidden)
{
	CGetTextHelper Options;
	return InitGetTextHelper(Options, bAll, bHeader, bHidden) ? GetCsvText(str, Options) : 0;
}

int CMyListCtrl::GetCsvText(CString& strCsv, const CGetTextHelper& Options) const
{
	ASSERT(Options.IsInitialized());

	strCsv.Empty();
	// Calculate estimated number of total characters. 
	// This does not include escaping of quote chars but assumes each cell is quoted.
	// With big lists, this avoids re-allocations.
	strCsv.Preallocate(
		Options.GetTextSizeLineFeed() +				// plain text size including line feeds
		Options.GetTotalColumnSeparatorCount() +	// column separator chars
		Options.GetTotalCells() * 2);				// assume all cells are quoted
	if (Options.m_bHeader)							// header is also requested
		GetCsvLine(strCsv, -1, Options);
	if (Options.m_bAll)								// all lines
	{
		for (int i = 0; i < Options.m_nItemCount; i++)
			GetCsvLine(strCsv, i, Options);
	}
	else											// selected lines only
	{
		POSITION pos = GetFirstSelectedItemPosition();
		while (pos)
			GetCsvLine(strCsv, GetNextSelectedItem(pos), Options);
	}
	return strCsv.GetLength();
}

// Append CSV formatted list row (or header line) to strLine.
void CMyListCtrl::GetCsvLine(CString& strLine, int nItem, const CGetTextHelper& Options) const
{
	bool bFirst = true;
	CString strItem;								// declare this outside the loop to avoid multiple allocations
	for (int i = 0; i < Options.m_nColumnCount; i++)
	{
		int j = Options.m_pOrder[i];				// Output according to the actual order on screen
		if (Options.m_pWidth[j])					// Hidden columns may be excluded
		{
			GetCellText(strItem, nItem, j, Options);	// Get text from header or sub-item
			// If the item contains the quote character it must be quoted
			//  and the quote characters must be esacped.
            int nPos = strItem.Find(m_cCsvQuote);
			bool bQuote = (nPos >= 0);
            while (nPos >= 0)						// escape quote chars
            {
                strItem.Insert(nPos, m_cCsvQuote);
                nPos = strItem.Find(m_cCsvQuote, nPos+2);
            }
			// Must quote if item contains separation char.
			// Must quote if item contains CR or LF.
			// May quote if item contains space or TAB (does not detect Unicode spaces).
			if (!bQuote)
			{
				if (strItem.Find(m_cCsvSep) >= 0 ||
					strItem.FindOneOf(m_bCsvQuoteSpace ? _T(" \t\r\n") : _T("\r\n")) >= 0)
				{
					bQuote = true;
				}
			}
			if (!bFirst)
				strLine += m_cCsvSep;				// separate columns
			if (bQuote)
				strLine += m_cCsvQuote;
			strLine += strItem;
			if (bQuote)
				strLine += m_cCsvQuote;
			bFirst = false;
		}
	}
	strLine += _T("\r\n");							// separate rows by new line
}

// Htmlify string.
LPCTSTR CMyListCtrl::TextToHtml(CString& str) const
{
	str.Replace(_T("&"), _T("&amp;"));				// this must be the first replacement!
	str.Replace(_T("\""), _T("&quot;"));
	str.Replace(_T("<"), _T("&lt;"));
	str.Replace(_T(">"), _T("&gt;"));
	return str.GetString();
}

// Get the content of all or selected lines into HTML table CString.
int CMyListCtrl::GetHtmlText(CString& str, bool bAll, bool bHeader, bool bHidden)
{
	CGetTextHelper Options;
	return InitGetTextHelper(Options, bAll, bHeader, bHidden) ? GetHtmlText(str, Options) : 0;
}

int CMyListCtrl::GetHtmlText(CString& strHtml, const CGetTextHelper& Options)
{
	ASSERT(Options.IsInitialized());

	CString strFontStyle;
	if (m_bHtmlFont)
		CHtmlFormat::GetFontStyle(strFontStyle, this);

	strHtml.Empty();
	// HTML: 
	//  "<table width="1" style="">\r\n</table>\r\n" once: 39
	//  "<thead>\r\n</thead>\r\n" once with header: 19
	//  "<tbody>\r\n</tbody>\r\n" once with header: 19
	//  "<tr></tr>" per line: 9
	//  "<td style="text-align:center"></td>" per cell using format "center" with max. length
	strHtml.Preallocate(
		Options.GetTextSizeLineFeed() +				// plain text size including line feeds
		Options.GetTotalCells() * 25 +				// per cell (td/th tags with style)
		Options.GetTotalCells() * 2 +				// some extra space for HTMLifying and nbsp for empty cells
		Options.GetTotalLines() * 9 +				// per line (tr tags)
		strFontStyle.GetLength() +					// per table (font style in table)
		(Options.m_bHeader ? 77 : 39));				// per table (table/thead/tbody tags)
	if (m_bHtmlFont)
	{
		strHtml.Format(_T("<table border=\"%d\" style=\"%s\">\r\n"), 
			m_nHtmlBorder, strFontStyle.GetString());
	}
	else
	{
		strHtml.Format(_T("<table border=\"%d\">\r\n"), 
			m_nHtmlBorder);
	}
	if (Options.m_bHeader)							// header is also requested
	{
		strHtml += _T("<thead>\r\n");
		GetHtmlLine(strHtml, -1, Options);
		strHtml += _T("</thead>\r\n");
		strHtml += _T("<tbody>\r\n");
	}
	if (Options.m_bAll)								// all lines
	{
		for (int i = 0; i < Options.m_nItemCount; i++)
			GetHtmlLine(strHtml, i, Options);
	}
	else											// selected lines only
	{
		POSITION pos = GetFirstSelectedItemPosition();
		while (pos)
			GetHtmlLine(strHtml, GetNextSelectedItem(pos), Options);
	}
	if (Options.m_bHeader)
		strHtml += _T("</tbody>\r\n");				// end of table body
	strHtml += _T("</table>\r\n");					// end of table
	return strHtml.GetLength();
}

// Append HTML formatted list row (or header line) to strLine.
void CMyListCtrl::GetHtmlLine(CString& strLine, int nItem, const CGetTextHelper& Options) const
{
	CString strItem;								// declare this outside the loop to avoid multiple allocations
	strLine += _T("<tr>");
	for (int i = 0; i < Options.m_nColumnCount; i++)
	{
		int j = Options.m_pOrder[i];				// Output according to the actual order on screen
		if (Options.m_pWidth[j])					// Hidden columns may be excluded
		{
			GetCellText(strItem, nItem, j, Options);	// Get text from header or sub-item
			if (!strItem.IsEmpty())
				TextToHtml(strItem);				// HTMLify text
			else // if (m_nHtmlBorder)
				strItem = _T("&nbsp;");				// ensures that empty cells are rendered with border
			int nFormat = Options.m_pFormat[j] & LVCFMT_JUSTIFYMASK;
			LPCTSTR lpszAlign = _T("left");
			if (LVCFMT_RIGHT == nFormat)
				lpszAlign = _T("right");
			else if (LVCFMT_CENTER == nFormat)
				lpszAlign = _T("center");
			strLine.AppendFormat(
//				_T("<t%c align=\"%s\">"), 
				_T("<t%c style=\"text-align:%s\">"), 
				nItem < 0 ? _T('h') : _T('d'), 
				lpszAlign);
			strLine += strItem;
			strLine += (nItem < 0) ? _T("</th>") : _T("</td>");
		}
	}
	strLine += _T("</tr>\r\n");						// end of row
}

// Get text for header item or sub-item.
void CMyListCtrl::GetCellText(CString& str, int nItem, int nColumn, const CGetTextHelper& Options) const
{
	if (nItem < 0)
		GetColumnTitle(str, nColumn);
	else
	{
		// If list has check boxes (always in column 0), add the check state.
		// Must also check the state image value (check box not shown if value is zero).
		if (0 == nColumn && Options.m_bHasCheckBoxes)
		{
			UINT nState = GetItemState(nItem, LVIS_STATEIMAGEMASK);
			if (INDEXTOSTATEIMAGEMASK(1) == nState)
				str = _T("[ ] ");
			else if (INDEXTOSTATEIMAGEMASK(2) == nState)
				str = _T("[X] ");
			else
				str = _T("    ");
			str += GetItemText(nItem, nColumn);
		}
		else
			str = GetItemText(nItem, nColumn);
	}
}

// Get column label string.
void CMyListCtrl::GetColumnTitle(CString& str, int nColumn) const
{
	TCHAR lpBuffer[256] = _T("");					// should be big enough
	LVCOLUMN lvc = { };
	lvc.mask = LVCF_TEXT;
	lvc.pszText = lpBuffer;
	lvc.cchTextMax = ARRAYSIZE(lpBuffer);
	VERIFY(GetColumn(nColumn, &lvc));				// asserts if string too long for buffer
	str = lpBuffer;
}

CString CMyListCtrl::GetColumnTitle(int nColumn) const
{
	TCHAR lpBuffer[256] = _T("");					// should be big enough
	LVCOLUMN lvc = { };
	lvc.mask = LVCF_TEXT;
	lvc.pszText = lpBuffer;
	lvc.cchTextMax = ARRAYSIZE(lpBuffer);
	VERIFY(GetColumn(nColumn, &lvc));				// assert if string too long for buffer
	return CString(lpBuffer);
}

// Get column format.
int CMyListCtrl::GetColumnFormat(int nColumn) const
{
	LVCOLUMN lvc = { };
	lvc.mask = LVCF_FMT;
	GetColumn(nColumn, &lvc);
	return lvc.fmt;
}

