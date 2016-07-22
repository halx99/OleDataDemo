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

// MyEdit.cpp : 
//  - Support Drag & Drop as source and target.

#include "stdafx.h"
#include "OleDataDemo.h"
#include "OleDataSourceEx.h"
#include "MyEdit.h"
#include ".\myedit.h"

#include <io.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CMyEdit

IMPLEMENT_DYNAMIC(CMyEdit, CEdit)
CMyEdit::CMyEdit()
{
	m_bDragging = false;							// Internal var; set when this is drag & drop source
	m_bOleMove = false;								// Drag & drop moving not allowed
	m_bOleDropSelf = false;							// Copy/move inside control not allowed
	m_bSelectDropped = true;						// Select dropped text
	m_bAllowDropDesc = false;						// Use drop description text (Vista and later)
	m_nDragImageType = DRAG_IMAGE_FROM_TEXT;		// Drag image type
	m_nDragImageScale = 100;						// Scaling value for drag images
	m_nDropLen = 0;									// Internal var; used with drop target
#ifdef MYEDIT_DROP_FILES							// HDROP handled by this class
	m_nDropFiles = -1;								// Accept file names from Explorer
#endif
}

CMyEdit::~CMyEdit()
{
}

BEGIN_MESSAGE_MAP(CMyEdit, CEdit)
	ON_WM_LBUTTONDOWN()								// OLE Drag&Drop source handling
	ON_WM_CONTEXTMENU()
	ON_COMMAND(ID_EDIT_CLEAR, OnEditClear)
	ON_COMMAND(ID_EDIT_COPY, OnEditCopy)
	ON_COMMAND(ID_EDIT_COPY_ALL, OnEditCopyAll)
	ON_COMMAND(ID_EDIT_CUT, OnEditCut)
//	ON_COMMAND(ID_EDIT_DESELECT, OnEditDeselect)
	ON_COMMAND(ID_EDIT_PASTE, OnEditPaste)
	ON_COMMAND(ID_EDIT_SELECT_ALL, OnEditSelectAll)
	ON_COMMAND(ID_EDIT_UNDO, OnEditUndo)
	ON_COMMAND(ID_PASTESPECIAL_CSV, OnEditPasteCsv)
	ON_COMMAND(ID_PASTESPECIAL_HTML, OnEditPasteHtml)
	ON_COMMAND(ID_PASTESPECIAL_HTML_ALL, OnEditPasteHtmlAll)
	ON_COMMAND(ID_PASTESPECIAL_RTF, OnEditPasteRtf)
	ON_COMMAND(ID_PASTESPECIAL_FILENAMES, OnEditPasteFileNames)
	ON_COMMAND(ID_PASTESPECIAL_FILECONTENT, OnEditPasteFileContent)
	ON_MESSAGE(WM_APP_DRAG_ENTER, OnDragEnter)		// drag callback messages
	ON_MESSAGE(WM_APP_DRAG_OVER, OnDragOver)
	ON_MESSAGE(WM_APP_DROP_EX, OnDropEx)
	ON_MESSAGE(WM_APP_DRAG_SCROLL, OnDragScroll)
	ON_MESSAGE(WM_APP_DO_SCROLL, OnDoScroll)
//	ON_MESSAGE(WM_APP_DRAG_LEAVE, OnDragLeave)
END_MESSAGE_MAP()

void CMyEdit::Init(LPCTSTR lpszText /*= NULL*/, UINT nLimit /*= 0*/)
{
	// Should use this when a drop target.
	// Unfortunately it can't be changed during run-time (e.g. upon OnDragEnter)
//	ASSERT(GetStyle() & ES_NOHIDESEL);
#ifdef MYEDIT_DROP_FILES							// HDROP handled by this class
	ModifyStyleEx(WS_EX_ACCEPTFILES, 0);			// Disable HDROP handling by CWnd
#endif

	// m_pDropTarget is CWnd member that is set to the address of the COleDropTarget class
	// when registering. If it is not NULL, this class has been already registered.
	if (NULL == m_pDropTarget)
    {
		// Use messages to pass drag events to this class.
		m_DropTarget.SetUseMsg();
		// Enable drag auto scrolling when over scroll bars or inset region.
		if (GetStyle() & ES_MULTILINE)
			m_DropTarget.SetScrollMode(AUTO_SCROLL_BARS | AUTO_SCROLL_INSETS);
		VERIFY(m_DropTarget.Register(this)); 
	}
	if (lpszText)
		SetWindowText(lpszText);
	if (nLimit)
		SetLimitText(nLimit);
}

#if 0
// Register this window as drop target upon creation.
// OnCreate() is not called for template created controls.
// So it should be done here.
void CMyEdit::PreSubclassWindow()
{
	Init();
	CEdit::PreSubclassWindow();
}
#endif

// Check if data can be pasted via OLE Drag & Drop.
// The length of the drop text is cached in m_nDropLen to avoid multiple checkings
//  when called repeatedly by OnDragOver.
// Perform these checks here:
// - Is the window enabled?
// - Is the control not read only?
// - Is a supported clipboard format available?
// - Does data fit (text must not contain line feeds if a single line edit control)?
LRESULT CMyEdit::OnDragEnter(WPARAM wParam, LPARAM lParam)
{
	m_nDropLen = 0;
	// Basic requirements:
	// - Window is enabled,
	// - Control is not read only,
	// - Dropping to ourself allowed or source is not this control.
	if (CanEdit() && (m_bOleDropSelf || !m_bDragging))
	{
		COleDataObject *pDataObject = reinterpret_cast<COleDataObject*>(wParam);
		// Text length in characters for Unicode, ANSI, and OEM text and number of lines.
		// Zero if neither CF_TEXT, CF_OEMTEXT, nor CF_UNICODETEXT data available.
		unsigned nLines = 0;
		m_nDropLen = static_cast<int>(m_DropTarget.GetCharCount(pDataObject, &nLines));
		if (m_nDropLen)
		{
			if (nLines > 1 && !(GetStyle() & ES_MULTILINE))
				m_nDropLen = 0;						// Can't drop multiline text on single line edit control
		}
#ifdef MYEDIT_DROP_FILES
		// NOTE:
		// When the window has the WS_EX_ACCEPTFILES style bit set
		//  this function is called but the return value is ignored (uses DROPEFFECT_COPY).
		else if (m_nDropFiles &&
//			!(GetExStyle() & WS_EX_ACCEPTFILES) &&
			pDataObject->IsDataAvailable(CF_HDROP))
		{
			m_nDropLen = static_cast<int>(m_DropTarget.GetFileListAsTextLength(pDataObject, &nLines));
			if (m_nDropFiles > 0 &&					// There is a limit for the number of file names
				nLines > static_cast<unsigned>(m_nDropFiles))
			{
				m_nDropLen = 0;						// Too many file names
			}
		}
#endif
	}
	// Call OnDragOver() to get and return the drop effect.
	return OnDragOver(wParam, lParam);
}

// Perform these checks here:
// - Determine the effect according to the key state
// - Does data fit (eg edit control would not reach its limit)?
// - Avoid dropping when over the source selection
LRESULT CMyEdit::OnDragOver(WPARAM wParam, LPARAM lParam)
{
	DROPEFFECT dropEffect = DROPEFFECT_NONE;
	if (m_nDropLen)
	{
		COleDropTargetEx::DragParameters *pParams = reinterpret_cast<COleDropTargetEx::DragParameters*>(lParam);
		CRect rcClient;								// Check if on the client area
		GetClientRect(rcClient);					// So scroll bars are excluded.
		if (rcClient.PtInRect(pParams->point))
		{
			// Get the max. number of chars that may be inserted without reaching the limit.
			int nAvail = GetLimitText();
#ifdef MYEDIT_DROP_FILES
			COleDataObject *pDataObject = reinterpret_cast<COleDataObject*>(wParam);
			if (pDataObject->IsDataAvailable(CF_HDROP))
			{
				dropEffect = DROPEFFECT_COPY;		// drag names, not files
			}
			else
#endif
			{
				if (m_bDragging && !m_bOleMove)		// Moving to ourself not allowed
					dropEffect = DROPEFFECT_COPY;
				else								// Move by default, copy when control key down.
					dropEffect = (pParams->dwKeyState & MK_CONTROL) ? DROPEFFECT_COPY : DROPEFFECT_MOVE;
			}
			CRect rect;
			bool bOnSel = GetSelectionRect(rect, &pParams->point);
			if (m_bDragging)						// source is this control
			{
				// When dropping on ourself, we should do nothing when over the selection:
				// - Moving (copy and delete) would not change the content.
				// - Copying may happen inadvertently when clicking on a selection and 
				//   releasing the mouse button shortly thereafter.
				if (bOnSel)							// do nothing when over the selection
					dropEffect = DROPEFFECT_NONE;
				else if (DROPEFFECT_COPY == dropEffect)
					nAvail -= GetWindowTextLength();
			}
			else if (!bOnSel)						// when dropping outside the selection,
				nAvail -= GetWindowTextLength();	//  adjust available space by content length
			if (m_nDropLen > nAvail)				// dropping would exceed the limit
				dropEffect = DROPEFFECT_NONE;
		}
	}
	return m_DropTarget.FilterDropEffect(dropEffect);
}

// Paste OLE drag & drop data.
// Dropping is handled by this function.
// So we don't need a OnDrop() function.
LRESULT CMyEdit::OnDropEx(WPARAM wParam, LPARAM lParam)
{
	// Update info list of main dialog.
	AfxGetMainWnd()->SendMessage(WM_USER_UPDATE_INFO, wParam, 0);

	DROPEFFECT dwEffect = DROPEFFECT_NONE;
	COleDropTargetEx::DragParameters *pParams = reinterpret_cast<COleDropTargetEx::DragParameters*>(lParam);
	if (pParams->dropEffect)
	{
		COleDataObject *pDataObject = reinterpret_cast<COleDataObject*>(wParam);
		// This will get an allocated copy of CF_TEXT, CF_OEMTEXT, or CF_UNICODETEXT data.
		// Text is converted to multi byte / Unicode according to project settings
		//  using CF_LOCALE if present.
		// It performs size checking so that invalid data without null terminator can be handled.
		LPCTSTR lpszText = m_DropTarget.GetText(pDataObject);
#ifdef MYEDIT_DROP_FILES
		// This may be also performed using the OnDropFiles() handler (WM_DROPFILES message).
		if (NULL == lpszText && 
			pDataObject->IsDataAvailable(CF_HDROP))
		{
			// This will get an allocated string with CR-LF terminated file names.
			lpszText = m_DropTarget.GetFileListAsText(pDataObject);
		}
#endif
		if (lpszText && *lpszText != _T('\0'))
		{
			BOOL bCanUndo = TRUE;
			bool bReplace = false;
			int nStart, nEnd;
			GetSel(nStart, nEnd);					// get current selection
			dwEffect = pParams->dropEffect;			// return default drop effect
			// Get character index from mouse point.
			// NOTE: With edit controls, the low word is always the character index
			//  relative to the control text, not to the beginning of line!
			int nPos = LOWORD(CharFromPos(pParams->point));
			if (nEnd > nStart)						// has a selection (always true when source is this control)
			{
				CRect rect;
				if (GetSelectionRect(rect, &pParams->point))
					bReplace = true;				// replace the current selection when over it
				// Moving text inside control:
				// - Delete selection
				// - Insert text at mouse position
				else if (m_bDragging && (pParams->dropEffect & DROPEFFECT_MOVE))
				{
					bCanUndo = FALSE;				// Can't undo this operation
					Clear();						// delete selection
					if (nPos >= nEnd)				// adjust position if right of deleted selection
						nPos -= nEnd - nStart;
					dwEffect = DROPEFFECT_COPY;		// return COPY effect; selection has just been deleted
				}
			}
			if (!bReplace)
				SetSel(nPos, nPos, TRUE);			// set caret pos (scrolling not necessary, pos is visible)
			ReplaceSel(lpszText, bCanUndo);			// replace selection or insert text at caret pos
			if (m_bSelectDropped)					// optional select the dropped text
			{
				if (bReplace)						// if current selection has been replaced
					nPos = nStart;					//  use the start position of the replaced selection
				SetSel(nPos, nPos + static_cast<int>(_tcslen(lpszText)), TRUE);
			}
		}
		delete [] lpszText;
	}
	m_nDropLen = 0;
	return dwEffect;
}

// Auto scrolling.
// COleDropTargetEx handles hits on scrollbars and when over the inset region.
LRESULT CMyEdit::OnDragScroll(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	// Use COleDropTragetEx default handling according to auto scroll flags.
	// Can't use LineScroll() with single line edit control.
	return (GetStyle() & ES_MULTILINE) ? DROPEFFECT_SCROLL : DROPEFFECT_NONE;
}

// When scrolling should be performed, the WM_APP_DO_SCROLL message is send by COleDropTargetEx.
LRESULT CMyEdit::OnDoScroll(WPARAM wParam, LPARAM lParam)
{
	LineScroll(static_cast<int>(wParam), static_cast<int>(lParam));
	return 0;
}

#if 0
int CMyEdit::GetCharWidth()
{
	if (0 == m_nCharWidth)
	{
		CDC* pDC = GetDC();
		CSize sz = pDC->GetTextExtent(_T("X"), 1);
		m_nCharWidth = sz.cx + 1;
	}
	return m_nCharWidth;
}

int CMyEdit::GetLineHeight()
{
	if (0 == m_nLineHeight)
	{
		CDC* pDC = GetDC();
		CSize sz = pDC->GetTextExtent(_T("Xq"), 2);
		m_nLineHeight = sz.cy + 1;
	}
	return m_nLineHeight;
}
#endif

// Get the visible rect for the current selection and/or check if pPoint is on the selection.
// rect is set to the visible part of the selection.
// With multi-line selections, left and right are set to the client rect.
// Returns false if nothing is selected.
// If pPoint is specified, true is returned only when the point is on the selection.
// NOTE: Function can't be const because we call GetDC().
// TODO: 
// - Check with scrolled content.
// - May cache char width and line height value in member vars.
//   Requires detection of font changing.
bool CMyEdit::GetSelectionRect(CRect& rect, const CPoint *pPoint /*= NULL*/)
{
	bool bRet = false;
	int nStart, nEnd;
	GetSel(nStart, nEnd);
	if (nEnd > nStart)
	{
		GetClientRect(&rect);
		CPoint pt1 = PosFromChar(nStart);			// position of first selected char
		CPoint pt2 = PosFromChar(nEnd - 1);			// position of last selected char
		CPoint pt3 = PosFromChar(nEnd);				// next char behind selection
		CDC* pDC = GetDC();
		int nLineHeight = 0;
		int nRight = rect.right;
		int nBottom = rect.bottom;
		// With single line edit controls, pt3 is invalid (negative positions)
		//  when the last char is included in the selection.
		if (pt3.x < 0 || pt3.y < 0)
		{
			CSize sz = pDC->GetTextExtent(_T("Xq"), 2);
			nLineHeight = sz.cy + 1;
			sz = pDC->GetTextExtent(_T("X"), 1);
			nRight = pt2.x + sz.cx;
			nBottom = pt2.y + nLineHeight;
		}
		else if (pt3.x > pt2.x)						// next char is right of the last selected char
		{
			CSize sz = pDC->GetTextExtent(_T("Xq"), 2);
			nLineHeight = sz.cy + 1;
			nRight = pt3.x - 1;
			nBottom = pt2.y + nLineHeight;
		}
		else										// pt3 is on next line
		{
			CSize sz = pDC->GetTextExtent(_T("X"), 1);
			nRight = pt2.x + sz.cx;
			nLineHeight = pt3.y - pt2.y;
			nBottom = pt3.y;
		}
		if (pt1.y > 0)								// Top position from first selected char.
			rect.top = pt1.y;
		if (nBottom < rect.bottom)					// Bottom position from last selected char and the char height.
			rect.bottom = nBottom;
		// With single line selections, the left and right positions are specified by the selection.
		// With multi line selections, left and right are from the client rect.
//		if (LineFromChar(nStart) == LineFromChar(nEnd - 1))
		if (pt1.y == pt2.y)							// single line
		{											//  left and right according to selection
			if (pt1.x > 0)
				rect.left = pt1.x;
			if (nRight < rect.right)
				rect.right = nRight;
		}
		if (pPoint)
		{
			bRet = (FALSE != rect.PtInRect(*pPoint));
			if (bRet) // && pt1.y != pt2.y)
			{
				// Check if left of first char when on first selected line
				// Check if right of last char when on last selected line
				if ((pPoint->y <= pt1.y + nLineHeight && pPoint->x < pt1.x) ||
					(pPoint->y >= pt2.y && pPoint->x > nRight))
					bRet = false;
			}
		}
		else
			bRet = true;
	}
	return bRet;
}

void CMyEdit::GetSelectedText(CString& str) const
{
	int nStart, nEnd;
	GetSel(nStart, nEnd);
	if (nEnd > nStart)
	{
		CString s;
		GetWindowText(s);
		str = s.Mid(nStart, nEnd - nStart);
	}
	else
		str = _T("");
}

// Cache data for clipboard and OLE Drag & Drop.
COleDataSourceEx* CMyEdit::CacheOleData(const CString& str, bool bClipboard)
{
	COleDataSourceEx *pDataSrc = new COleDataSourceEx(this);
	if (pDataSrc)
	{
		// Cache standard text formats.
		pDataSrc->CacheText(str.GetString(), bClipboard);
		// Cache text as CF_BITMAP
		pDataSrc->CacheBitmapFromText(str.GetString(), this);
		if (bClipboard)
		{
			pDataSrc->SetClipboard();
			pDataSrc = NULL;						// object has been deleted by SetClipboard()
		}
	}
	return pDataSrc;
}

// OLE Drag & Drop source.
//
// When starting a drag & drop operation (calling DoDragDrop), all mouse events are handled by OLE
//  until the mouse button is released or the ESC key is pressed.
// Therefore, we must not pass the LBUTTONDOWN message to the default handler when starting drag here.
void CMyEdit::OnLButtonDown(UINT nFlags, CPoint point)
{
	m_bDragging = false;
	bool bHandled = false;							// set when message processed by dragging
	bool bSendButtonUpMsg = false;					// set when a single click should be simulated
	// If the selection is always shown (ES_NOHIDESEL style set), we may
	//  start dragging even when the control does not has the focus.
	if (GetFocus() == this || (GetStyle() & ES_NOHIDESEL))
	{
		CRect rect;
		if (GetSelectionRect(rect, &point))			// clicked on the selection
		{
			CString strText;
			GetSelectedText(strText);

			COleDataSourceEx *pDataSrc = CacheOleData(strText, false);

			// This must be called before setting the drag image.
			if (m_bAllowDropDesc)
				pDataSrc->AllowDropDescriptionText();

			CPoint pt(-1, 30000);					// Cursor is centered below the image
			switch (m_nDragImageType & DRAG_IMAGE_FROM_MASK)
			{
			case DRAG_IMAGE_FROM_TEXT :
				// Create drag image from text string.
				pDataSrc->InitDragImage(this, strText.GetString(), NULL, 0 != (m_nDragImageType & DRAG_IMAGE_HITEXT));
				break;
			case DRAG_IMAGE_FROM_RES :
				// Use bitmap from resources as drag image
				pDataSrc->InitDragImage(IDB_BITMAP_DRAGME, &pt, CLR_INVALID);
				break;
			case DRAG_IMAGE_FROM_HWND :
				// This window handles DI_GETDRAGIMAGE messages.
//				pDataSrc->SetDragImageWindow(m_hWnd, &point);
//				break;
			case DRAG_IMAGE_FROM_SYSTEM :
				// Let Windows create the drag image.
//				pDataSrc->SetDragImageWindow(NULL, &point);
				pDataSrc->SetDragImageWindow(NULL, NULL);
				break;
			case DRAG_IMAGE_FROM_EXT :
				// Get drag image from file extension using the file type icon
				pDataSrc->InitDragImage(_T(".txt"), m_nDragImageScale, &pt);
				break;
			case DRAG_IMAGE_FROM_BITMAP :
				// Create a bitmap from text using the font and colors of this window.
				{
					CBitmap Bitmap;
					if (pDataSrc->BitmapFromText(Bitmap, this, strText.GetString()))
					{
						CDC *pDC = GetDC();
						pDataSrc->SetDragImage(Bitmap, &pt, pDC ? pDC->GetBkColor() : (COLORREF)-1);
//						pDataSrc->InitDragImage((HBITMAP)Bitmap, m_nDragImageScale, &pt, pDC ? pDC->GetBkColor() : (COLORREF)-1);
					}
					else
						pDataSrc->SetDragImageWindow(NULL, NULL);
				}
				break;
			// case DRAG_IMAGE_FROM_CAPT :
			default :
				// Get drag image by capturing content from screen with optional scaling.
				pDataSrc->InitDragImage(this, &rect, m_nDragImageScale);
			}

			// NOTE:
			//  Data is moved by default and copied when the control key is down.
			//  This behaviour may be not expected by many users with edit controls.
			DROPEFFECT nEffect = DROPEFFECT_COPY;
			if (m_bOleMove && CanEdit())			// moving allowed, not read only, window is enabled
				nEffect |= DROPEFFECT_MOVE;
			m_bDragging = true;						// dragging will be started
			if (DROPEFFECT_MOVE == pDataSrc->DoDragDropEx(nEffect))
				Clear();							// clear selection if data has been moved
			m_bDragging = false;
			// If dragging has not been started because the mouse button was released
			//  while inside the start rect and before the delay time has expired, we
			//  can simulate a single click. With edit controls, this will set the focus to
			//  the control if it does not has the focus already, remove the current selection, and
			//  set the cursor caret at the clicked position.
			bHandled = true;
			if (DRAG_RES_RELEASED == pDataSrc->GetDragResult())
				bSendButtonUpMsg = true;
			pDataSrc->InternalRelease();
		}
	}
	if (!bHandled || bSendButtonUpMsg)
	{
		CEdit::OnLButtonDown(nFlags, point);		// default handling of left mouse button down
		if (bSendButtonUpMsg)						// simulate single click
			PostMessage(WM_LBUTTONUP, nFlags, MAKELPARAM(point.x, point.y));
	}
}

bool CMyEdit::CanSelectAll() const
{
	int nStart, nEnd;
	GetSel(nStart, nEnd);
	return (nStart > 0 || nEnd < GetWindowTextLength());
}

bool CMyEdit::HasSelection() const
{
	int nStart, nEnd;
	GetSel(nStart, nEnd);
	return (nEnd > nStart);
//	return (nStart != nEnd && GetWindowTextLength() != 0);
}

// Command handlers.
void CMyEdit::OnEditCopy()
{
	// The Copy() function provides CF_TEXT, CF_UNICODETEXT, CF_OEM_TEXT, and CF_LOCALE
	//  as HGLOBAL and IStream.
//	Copy();
	// With CacheOleData(), a bitmap will be also provided.
	CString strText;
	GetSelectedText(strText);
	CacheOleData(strText, true);
}

void CMyEdit::OnEditCopyAll()
{
	CString strText;
	GetWindowText(strText);
	CacheOleData(strText, true);
}

void CMyEdit::OnEditSelectAll()
{
	SetSel(0, -1, 1);
}

void CMyEdit::OnEditDeselect()
{
	SetSel(-1, 0, 1);
}

void CMyEdit::OnEditUndo()
{
	Undo();
}

void CMyEdit::OnEditPaste()
{
#if 1
	Paste();
#else
	// This is just to demonstrate how clipboard data can be pasted.
    COleDataObject DataObj;
    DataObj.AttachClipboard();
	LPCTSTR lpszText = m_DropTarget.GetText(&DataObj);
    if (lpszText)
    {
		ReplaceSel(lpszText, 1);
		delete [] lpszText;
    }
	DataObj.Release();
#endif
}

// Paste CSV text as plain text.
void CMyEdit::OnEditPasteCsv()
{
	PasteAnsiText(CF_CSV_STR);
}

// Paste HTML text as plain text (complete HTML document if present).
void CMyEdit::OnEditPasteHtmlAll()
{
	PasteHtml(true);
}

// Paste HTML text as plain text (selected parts only).
void CMyEdit::OnEditPasteHtml()
{
	PasteHtml(false);
}

void CMyEdit::PasteHtml(bool bAll)
{
	COleDataObject DataObj;
	DataObj.AttachClipboard();
	// Clipboard HTML format is UTF-8 encoded.
	// This retrieves a converted string.
	LPCTSTR lpszText = m_DropTarget.GetHtmlString(&DataObj, bAll);
	if (lpszText)
	{
		ReplaceSel(lpszText, 1);
		delete [] lpszText;
	}
	DataObj.Release();
}

// Paste RTF text as plain text including format options.
// RTF is ASCII text (code page 20127).
void CMyEdit::OnEditPasteRtf()
{
	PasteAnsiText(CF_RTF_STR, 20127);
}

#if 0
void CMyEdit::OnEditPasteRtfAsText()
{
	PasteAnsiText(_T("RTF as Text"), 20127);
}

void CMyEdit::OnEditPasteRtfWithoutObjects()
{
	PasteAnsiText(_T("Rich Text Format Without Objects"), 20127);
}
#endif

// Paste plain text from custom formats.
void CMyEdit::PasteAnsiText(CLIPFORMAT cfFormat, UINT nCP /*= CP_ACP*/)
{
	COleDataObject DataObj;
	DataObj.AttachClipboard();
	if (DataObj.IsDataAvailable(cfFormat))
	{
		LPCTSTR lpszText = m_DropTarget.GetString(cfFormat, &DataObj, nCP);
		if (lpszText)
		{
			ReplaceSel(lpszText, 1);
			delete [] lpszText;
		}
	}
	DataObj.Release();
}

// Paste file names from CF_HDROP.
void CMyEdit::OnEditPasteFileNames()
{
	COleDataObject DataObj;
	DataObj.AttachClipboard();
	if (DataObj.IsDataAvailable(CF_HDROP))
	{
		// Get CR-LF separated list of file names into allocated buffer
		LPTSTR lpszText = m_DropTarget.GetFileListAsText(&DataObj);
		if (lpszText)
		{
			ReplaceSel(lpszText, 1);
			delete [] lpszText;
		}
	}
	DataObj.Release();
}

// Paste file content.
// NOTES:
// - Does not check if file is a text file
void CMyEdit::OnEditPasteFileContent()
{
	COleDataObject DataObj;
	DataObj.AttachClipboard();
	CString strFileName = m_DropTarget.GetSingleFileName(&DataObj);
	if (!strFileName.IsEmpty())
	{
		FILE *f;
		if (0 == _tfopen_s(&f, strFileName.GetString(), _T("rb")))
		{
			LPTSTR lpszText = NULL;
			long nSize = _filelength(_fileno(f));
			if (nSize)
			{
				// Add space for trailing null byte if content can be used without conversion.
				LPBYTE pBuf = new BYTE[nSize + sizeof(TCHAR)];
				fread(pBuf, 1, nSize, f);
				// Check for UTF-16LE and UTF-8 BOM.
				UINT nCP = CP_ACP;
				if (0xFF == pBuf[0] && 0xFE == pBuf[1])
					nCP = 1200;
				else if (0xEF == pBuf[0] && 0xBB == pBuf[0] && 0xBF == pBuf[0])
					nCP = CP_UTF8;
#ifdef _UNICODE
				if (1200 == nCP)
				{
					lpszText = reinterpret_cast<LPTSTR>(pBuf);
					lpszText[nSize / sizeof(WCHAR)] = _T('\0');
					pBuf = NULL;
				}
				else
					lpszText = CDragDropHelper::MultiByteToWideChar(reinterpret_cast<LPCSTR>(pBuf), nSize, nCP);
#else
				if (1200 == nCP)
					lpszText = CDragDropHelper::WideCharToMultiByte(reinterpret_cast<LPCWSTR>(pBuf), nSize / sizeof(WCHAR), nCP);
				else if (CP_UTF8 == nCP)
					lpszText = CDragDropHelper::MultiByteToMultiByte(reinterpret_cast<LPCSTR>(pBuf), nSize, nCP);
				else
				{
					lpszText = reinterpret_cast<LPSTR>(pBuf);
					lpszText[nSize] = _T('\0');
					pBuf = NULL;
				}
#endif
				delete [] pBuf;
			}
			fclose(f);
			if (lpszText)
			{
				ReplaceSel(lpszText, 1);
				delete [] lpszText;
			}
		}
	}
	DataObj.Release();
}

void CMyEdit::OnEditCut()
{
	Cut();
}

void CMyEdit::OnEditClear()
{
	int nStart, nEnd;
	GetSel(nStart, nEnd);
	if (nStart == nEnd && 							// if nothing selected and not at end
		nStart < GetWindowTextLength())
		SetSel(nStart, nStart + 1);					//  select char right of cursor and delete it
	Clear();
}

#ifdef IDR_EDIT_MENU
void CMyEdit::OnContextMenu(CWnd* pWnd, CPoint point)
{
	int nStart, nEnd;
	GetSel(nStart, nEnd);
	int nLen = GetWindowTextLength();

	if (point.x < 0 && point.y < 0)					// if opened by keyboard (not mouse)
	{
		RECT rect;
		GetWindowRect(&rect);
		if (GetStyle() & ES_MULTILINE)
		{
			// Left top position of last char
			point = PosFromChar(nEnd >= nLen ? nLen - 1 : nEnd);
			// If not on last line, get y position from next line index.
			int nNext = LineIndex(1 + LineFromChar());
			if (nNext >= 0)
				point.y = PosFromChar(nNext).y;
			else
				point.y += 16;						// use a common line height
			ClientToScreen(&point);
			if (point.x < rect.left)
				point.x = rect.left;
			if (point.y > rect.bottom)
				point.y = rect.bottom;
		}
		else										// single line edit control
			point.SetPoint(rect.left, rect.bottom);	// below the control so that the text is not hidden
	}
	else 
	{
		POINT pt = point;
		ScreenToClient(&pt);						// Get point relative to edit control.
		CRect rect;									// Check if clicked inside edit area.
		GetClientRect(&rect);						// Must use client rect to exclude scroll bars.
		if (!rect.PtInRect(pt) ||					// clicked outside the edit area (e.g. on scroll bar)
			pWnd != static_cast<CWnd*>(this))
		{
			CEdit::OnContextMenu(pWnd, point);		// call default handler
			return;
		}
	}

	CMenu menu;
	VERIFY(menu.LoadMenu(IDR_EDIT_MENU));			// Load the menu.
	CMenu *pSub = menu.GetSubMenu(0);				// Get the pop-up menu.
	ASSERT(pSub != NULL);

	// When right clicking on the edit control we should get the focus first
	//  to show the active selection. 
	// Because many context commands use the selection (Copy, Cut, Clear, Deselect).
	if (GetFocus() != this && !(GetStyle() & ES_NOHIDESEL))
		SetFocus();

	// The Update UI mechanism does only work with frame window derived classes
	//  because the code that calls the UI handlers is not found in other classes.
	// See also http://support.microsoft.com/?scid=kb%3Ben-us%3B139469&x=10&y=12.
	// So we must do the job here.

	// Copy and select operations
	if (nLen == 0)									// edit control is empty
	{
		VERIFY(pSub->EnableMenuItem(ID_EDIT_COPY_ALL, MF_GRAYED) != -1);
	}
	if (nStart == nEnd || nLen == 0)				// nothing selected or edit control empty
	{
		VERIFY(pSub->EnableMenuItem(ID_EDIT_COPY, MF_GRAYED) != -1);
//		VERIFY(pSub->EnableMenuItem(ID_EDIT_DESELECT, MF_GRAYED) != -1);
	}
	if (nStart == 0 && nEnd >= nLen)				// all selected
		VERIFY(pSub->EnableMenuItem(ID_EDIT_SELECT_ALL, MF_GRAYED) != -1);

	// Edit operations
	bool bCanEdit = CanEdit();
	if (!bCanEdit || !CanUndo())
		VERIFY(pSub->EnableMenuItem(ID_EDIT_UNDO, MF_GRAYED) != -1);
	if (!bCanEdit || !CanPaste())
		VERIFY(pSub->EnableMenuItem(ID_EDIT_PASTE, MF_GRAYED) != -1);
	if (!bCanEdit || !CanCut())
		VERIFY(pSub->EnableMenuItem(ID_EDIT_CUT, MF_GRAYED) != -1);
	if (!bCanEdit || !CanClear())
		VERIFY(pSub->EnableMenuItem(ID_EDIT_CLEAR, MF_GRAYED) != -1);

	// Paste special sub-menu items
	CMenu * pSpecialMenu = pSub->GetSubMenu(5);
	ASSERT(pSpecialMenu != NULL);					// IDR_EDIT_MENU resource changed?
	UINT nFormat = ::RegisterClipboardFormat(CF_CSV_STR);
	if (!bCanEdit || nFormat == 0 || !::IsClipboardFormatAvailable(nFormat))
		pSpecialMenu->EnableMenuItem(ID_PASTESPECIAL_CSV, MF_GRAYED);
	nFormat = ::RegisterClipboardFormat(CF_HTML_STR);
	if (!bCanEdit || nFormat == 0 || !::IsClipboardFormatAvailable(nFormat))
	{
		pSpecialMenu->EnableMenuItem(ID_PASTESPECIAL_HTML, MF_GRAYED);
		pSpecialMenu->EnableMenuItem(ID_PASTESPECIAL_HTML_ALL, MF_GRAYED);
	}
	nFormat = ::RegisterClipboardFormat(CF_RTF_STR);
	if (!bCanEdit || nFormat == 0 || !::IsClipboardFormatAvailable(nFormat))
		pSpecialMenu->EnableMenuItem(ID_PASTESPECIAL_RTF, MF_GRAYED);
	if (!bCanEdit || !::IsClipboardFormatAvailable(CF_HDROP))
	{
		pSpecialMenu->EnableMenuItem(ID_PASTESPECIAL_FILENAMES, MF_GRAYED);
		pSpecialMenu->EnableMenuItem(ID_PASTESPECIAL_FILECONTENT, MF_GRAYED);
	}
	else
	{
		// Disable paste file content if more than one file name.
		COleDataObject DataObj;
		DataObj.AttachClipboard();
		if (1 != m_DropTarget.GetFileNameCount(&DataObj))
			pSpecialMenu->EnableMenuItem(ID_PASTESPECIAL_FILECONTENT, MF_GRAYED);
	}

	int nCmd = pSub->TrackPopupMenuEx(				// Display the popup menu and return the command.
		TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_VERPOSANIMATION | TPM_RETURNCMD | TPM_NONOTIFY,
		point.x, point.y, AfxGetMainWnd(), NULL);
	if (nCmd)
		SendMessage(WM_COMMAND, nCmd);
}
#endif
