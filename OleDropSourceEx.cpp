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

// COleDropSource derived class to support DropDescriptions

#include "StdAfx.h"
#include "OleDropSourceEx.h"

#include <basetsd.h>								// ULongToHandle() macro

// Must set this definition to include OEM constants from winuser.h
#ifndef OCR_NORMAL
#error Add '#define OEMRESOURCE' on top of stdafx.h
#endif

#ifndef DDWM_UPDATEWINDOW
#define DDWM_UPDATEWINDOW	(WM_USER + 3)
#endif
// This message is undocumented.
// If no drop description data object exists, the WPARAM
//  parameter specifies the cursor type and the system default text is shown
#define DDWM_SETCURSOR		(WM_USER + 2)

COleDropSourceEx::COleDropSourceEx()
{
	m_bSetCursor = true;							// Must set default cursor
	m_nResult = DRAG_RES_NOT_STARTED;				// Drag result flags
	m_pIDataObj = NULL;								// IDataObject of COleDataSourceEx
	m_pDropDescription = NULL;						// Drop description text of COleDataSourceEx
}

// Set the drag image cursor and text according to the drop effect and redraw.
// Cursor and description text (if enabled) are shown according to the passed drop effect
// when dwEffect is valid and a DropDescription data object does not exist or has the 
// DROPIMAGE_INVALID image type set. 
//
//  DDWM_SETCURSOR (WM_USER + 2)
//   WPARAM:  Cursor type
//    0 = Get type and text from DropDescription data object (behaviour is identical to DDWM_UPDATEWINDOW message)
//    1 = Can't drop (Stop sign)
//    2 = Move (Arrow)
//    3 = Copy (Plus sign)
//    4 = Link (Curved arrow)
//   LPARAM: 0
//  Set the drag image cursor and use the associated default text if no drop description
//   present or it has the invalid image type.
//  NOTE: When a drop description object is present and has not the DROPIMAGE_INVALID type
//   set, it is used regardless of the WPARAM value.
bool COleDropSourceEx::SetDragImageCursor(DROPEFFECT dwEffect)
{
	// Stored data is always a DWORD even with 64-bit apps.
	HWND hWnd = (HWND)ULongToHandle(CDragDropHelper::GetGlobalDataDWord(m_pIDataObj, _T("DragWindow")));
	if (hWnd)
	{
		WPARAM wParam = 0;								// Use DropDescription to get type and optional text
		switch (dwEffect & ~DROPEFFECT_SCROLL)
		{
		case DROPEFFECT_NONE : wParam = 1; break;
		case DROPEFFECT_COPY : wParam = 3; break;		// The order is not as for DROPEEFECT values!
		case DROPEFFECT_MOVE : wParam = 2; break;
		case DROPEFFECT_LINK : wParam = 4; break;
		}
		::SendMessage(hWnd, DDWM_SETCURSOR, wParam, 0);
	}
	return NULL != hWnd;
}

// Called by COleDataSource::DoDragDrop().
// Override used here to set result flags.
BOOL COleDropSourceEx::OnBeginDrag(CWnd* pWnd)
{
	// Base class version returns m_bDragStarted.
	BOOL bRet = COleDropSource::OnBeginDrag(pWnd);
	// Set result flags.
	// DRAG_RES_STARTED: Drag has been started by leaving the start rect or after delay time has expired.
	if (m_bDragStarted)
		m_nResult = DRAG_RES_STARTED;
	// DRAG_RES_RELEASED: Mouse button has been released before leaving start rect or delay time has expired.
	else if (GetKeyState((m_dwButtonDrop & MK_LBUTTON) ? VK_LBUTTON : VK_RBUTTON) >= 0)
		m_nResult = DRAG_RES_RELEASED;
	// DRAG_RES_CANCELLED: ESC pressed or other mouse button activated.
	else
		m_nResult = DRAG_RES_CANCELLED;
	return bRet;
}

// Override used here to set result flags.
SCODE COleDropSourceEx::QueryContinueDrag(BOOL bEscapePressed, DWORD dwKeyState)
{
	SCODE scRet = COleDropSource::QueryContinueDrag(bEscapePressed, dwKeyState);
	if (0 == (dwKeyState & m_dwButtonDrop))			// mouse button has been released
		m_nResult |= DRAG_RES_RELEASED;
	if (DRAGDROP_S_CANCEL == scRet)					// ESC pressed, other mouse button activated, or
		m_nResult |= DRAG_RES_CANCELLED;			//  mouse button released when not started
	return scRet;
}

// Handle feedback.
// This is called when returning from the OnDrag... functions of the drop target interface.
// When using drop descriptions, we must update the drag image window here to show
//  the new style cursors and the optional description text.
SCODE COleDropSourceEx::GiveFeedback(DROPEFFECT dropEffect)
{
	SCODE sc = COleDropSource::GiveFeedback(dropEffect);
	// Drag must be started, Windows version must be Vista or later
	//  and visual styles must be enabled.
    if (m_bDragStarted && m_pIDataObj)
    {
		// IsShowingLayered is true when the target window shows the drag image
		// (target window calls the IDropTargetHelper handler functions).
		bool bOldStyle = (0 == CDragDropHelper::GetGlobalDataDWord(m_pIDataObj, _T("IsShowingLayered")));
	
		// Check if the drop description data object must be updated:
		// - When entering a window that does not show drag images while the previous
		//   target has shown the image, the drop description should be set to invalid.
		//   Condition: bOldStyle && !m_bSetCursor
		// - When setting drop description text here is enabled and the target
		//   shows the drag image, the description should be updated if not done by the target.
		//   Condition: m_pDropDescription && !bOldStyle
		if ((bOldStyle && !m_bSetCursor) || (m_pDropDescription && !bOldStyle))
		{
			// Get DropDescription data if present.
			// If not present, cursor and optional text are shown like within the Windows Explorer.
			FORMATETC FormatEtc;
			STGMEDIUM StgMedium;
			if (CDragDropHelper::GetGlobalData(m_pIDataObj, CFSTR_DROPDESCRIPTION, FormatEtc, StgMedium))
			{
				bool bChangeDescription = false;		// Set when DropDescription will be changed here
				DROPDESCRIPTION *pDropDescription = static_cast<DROPDESCRIPTION*>(::GlobalLock(StgMedium.hGlobal));
				if (bOldStyle)
				{
					bChangeDescription = CDragDropHelper::ClearDescription(pDropDescription);
				}
				// The drop target is showing the drag image and new cursors and may have changed the description.
				// Drop targets that change the description should clear it when leaving the window
				// (set the image type to DROPIMAGE_INVALID from within OnDragLeave()).
				// The target may have set a special image type (LABEL, WARNING, NOIMAGE).
				// Then use it and do not change anything here.
				else if (pDropDescription->type <= DROPIMAGE_LINK)
				{
					DROPIMAGETYPE nImageType = CDragDropHelper::DropEffectToDropImage(dropEffect);
					// When the target has changed the description, it usually has set the correct type.
					// If the image type does not match the drop effect, set the text here.
					if (DROPIMAGE_INVALID != nImageType &&
						pDropDescription->type != nImageType)
					{
						// When text is available, set drop description image type and text.
						if (m_pDropDescription->HasText(nImageType))
						{
							bChangeDescription = true;
							pDropDescription->type = nImageType;
							m_pDropDescription->SetDescription(pDropDescription, nImageType);
						}
						// Set image type to invalid. Cursor is shown according to drop effect
						//  using system default text.
						else
							bChangeDescription = CDragDropHelper::ClearDescription(pDropDescription);
					}
				}	
				::GlobalUnlock(StgMedium.hGlobal);
				if (bChangeDescription)				// Update the DropDescription data object when 
				{									//  image type or description text has been changed here.
					if (FAILED(m_pIDataObj->SetData(&FormatEtc, &StgMedium, TRUE)))
						bChangeDescription = false;
				}
				if (!bChangeDescription)			// Description not changed or setting it failed
					::ReleaseStgMedium(&StgMedium);
			}
		}
		if (!bOldStyle)								// Show new style drag cursors.
		{
			if (m_bSetCursor)
			{
				// Set the default cursor.
				// Otherwise, the old style drag cursor remains when leaving a window
				//  that does not support drag images.
				// NOTE: 
				//  Add '#define OEMRESOURCE' on top of stdafx.h.
				//  This is necessary to include the OCR_ definitions from winuser.h.
				HCURSOR hCursor = (HCURSOR)LoadImage(
					NULL,							// hInst must be NULL for OEM images
					MAKEINTRESOURCE(OCR_NORMAL),	// default cursor
					IMAGE_CURSOR,					// image type is cursor
					0, 0,							// use system metrics values with LR_DEFAULTSIZE
					LR_DEFAULTSIZE | LR_SHARED);
				::SetCursor(hCursor);
			}
			// Update of the cursor may be suppressed if "IsShowingText" is false.
			// But this results in invisible cursors for very short moments.
//			if (0 != CDragDropHelper::GetGlobalDataDWord(m_pIDataObj, _T("IsShowingText")))
			// If a drop description object exists and the image type is not invalid,
			//  it is used. Otherwise use the default cursor and text according to the drop effect. 
			SetDragImageCursor(dropEffect);
			sc = S_OK;								// Don't show default (old style) drag cursor
        }
		// When the old style drag cursor is actually used, the Windows cursor must be set
		//  to the default arrow cursor when showing new style drag cursors the next time.
		// If a new style drag cursor is actually shown, the cursor has been set above.
		// By using this flag, the cursor must be set only once when entering a window that
		//  supports new style drag cursors and not with each call when moving over it.
		m_bSetCursor = bOldStyle;
    }
	return sc;
}

#ifdef _DEBUG
void COleDropSourceEx::Dump(CDumpContext& dc) const
{
	COleDropSource::Dump(dc);

	dc << "m_bSetCursor = " << m_bSetCursor;
	dc << "\nm_nResult = " << m_nResult;

	dc << "\n";
}
#endif //_DEBUG

