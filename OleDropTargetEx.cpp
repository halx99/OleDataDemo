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

#include "stdafx.h"
#include "OleDropTargetEx.h"

// Include header files for supported CWnd derived classes here.
//#include "MyClass.h"

/////////////////////////////////////////////////////////////////////////////
// COleDropTargetEx - OLE Drag & Drop target class supporting drag images
//  and drop descriptions.
//

// To show drag images when dragging content over your app:
//  Add a member of this class to the main frame class and call Register() from OnCreate().
//  With dialog applications, do this for your main dialog and call Register() from OnInitDialog().
//
// To support drop for CView derived windows:
// - Add a member of this class to the view class.
// - Register the drop target from within OnInitialUpdate().
// - Override the virtual handler functions in the view class.
//
// To support drop for other windows:
//  Method 1 (using callback functions)
//  - Add a member of this class to the window class.
//  - Add static handler functions to the window class.
//    The static functions may handle the events or call a non-static version.
//  - For those events that aren't handled by static functions,
//    add the non-static versions.
//  - Pass pointers to the static wrapper functions to this class.
//  - Register the drop target using OnInitDialog of the parent dialog
//    (optional using a provided Init function).
//  Method 2 (using user defined messages)
//  - Add a member of this class to the window class.
//  - Add standard message handlers for the user defined messages to the window class.
//    The WMP_APP_DROP_EX handler must be always added (return -1 if not used).
//    Other handlers are optional.
//  - Enable usage of messages by calling SetUseMsg(true)
//  - Register the drop target using OnInitDialog of the parent dialog
//    (optional using a provided Init function).
//  Method 3 (like MFC for CView)
//  - Add a member of this class to the window class.
//  - Add the handler functions to the window class.
//  - Register the drop target using OnInitDialog of the parent dialog
//    (optional using a provided Init function).
//  - Include window class header file here.
//  - Add code here to call the handlers.
//
//  Alternative option:
//  - Add handler code here and call it if the class matches (e.g. CEdit)
//  - Still requires registering and that this class is a member of the control class
//
//  Using auto scrolling
//   This class supports detection of auto scroll events when over scrollbars or on an inset region.
//   Which events should be handled is specified by flags. The window class must only
//   provide a callback function or accepting a user defined message to perform scrolling.
//
//   Method 1 (using callback function)
//   - Add a static handler function to the window class.
//     The static function may perform the scrolling or call a non-static version.
//   - Call SetAutoScrollMode() passing the desired flags and the address of the callback function.
//   Method 2 (using user defined message)
//   - Add a standard message handler for WM_APP_DO_SCROLL to the window class.
//   - Call SetAutoScrollMode() passing the desired flags.
//
// CRichEditCtrl
//  Rich edit controls already support drag & drop. But without drag images.
//  Using this class (or COleDropTarget) with rich edit controls is not possible
//  (registering the control window fails).

COleDropTargetEx::COleDropTargetEx()
{
	m_bUseMsg = false;								// Set by window class when using messages
	m_bUseDropDescription = false;					// If true, drop descriptions are updated by this class
	m_nScrollMode = 0;								// Set by window class to configure auto scrolling
	m_pOnDragEnter = NULL;							// Drop target callback functions
	m_pOnDragOver = NULL;
	m_pOnDropEx = NULL;
	m_pOnDrop = NULL;
	m_pOnDragLeave = NULL;
	m_pOnDragScroll = NULL;
	m_pScrollFunc = NULL;							// Scroll callback function
	m_pDropTargetHelper = NULL;						// IDropTargetHelper to show drag images
#if defined(IID_PPV_ARGS) // The IID_PPV_ARGS macro has been introduced with Visual Studio 2005
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pDropTargetHelper));
#else
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER,
		IID_IDropTargetHelper, reinterpret_cast<LPVOID*>(&m_pDropTargetHelper));
#endif
	m_bCanShowDescription = (m_pDropTargetHelper != NULL);	// Vista or later (drop descriptions are supported)
	if (m_bCanShowDescription)
	{
#if (WINVER < 0x0600)
		OSVERSIONINFOEX VerInfo;
		DWORDLONG dwlConditionMask = 0;
		::ZeroMemory(&VerInfo, sizeof(OSVERSIONINFOEX));
		VerInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
		VerInfo.dwMajorVersion = 6;
		VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_GREATER_EQUAL);
		m_bCanShowDescription = ::VerifyVersionInfo(&VerInfo, VER_MAJORVERSION, dwlConditionMask) && ::IsAppThemed();
#else
		m_bCanShowDescription = (0 != ::IsAppThemed());
#endif
	}
	ClearStateFlags();
}

COleDropTargetEx::~COleDropTargetEx()
{
	if (NULL != m_pDropTargetHelper)
		m_pDropTargetHelper->Release();					// release the IDropTargetHelper
}

// Clear drag session related flags.
void COleDropTargetEx::ClearStateFlags()
{
	m_bDescriptionUpdated = false;					// Internal flag set when drop description has been updated
	m_bEntered = false;								// Set when entered the target. Used by OnDragScroll().
	m_bHasDragImage = false;						// Set when source provides a drag image.
	m_bTextAllowed = false;							// Set when source has enabled drop description text.
	m_nPreferredDropEffect = DROPEFFECT_NONE;		// Cached CFSTR_PREFERREDDROPEFFECT value.
}

// Set default drop description text string.
// Call this for each image type that should show a specific description.
// All others will use the system default text.
bool COleDropTargetEx::SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert /*= NULL*/)
{
	bool bRet = false;
	if (m_bCanShowDescription)						// Vista or later and themes are enabled
	{
		bRet = m_DropDescription.SetText(nType, lpszText, lpszText1);
		if (bRet && lpszInsert)
			m_DropDescription.SetInsert(lpszInsert);
		m_bUseDropDescription |= bRet;				// drop description text is available
	}
	return bRet;
}

// Set default drop description insert string.
bool COleDropTargetEx::SetDropInsertText(LPCWSTR lpszText)
{
	return m_bCanShowDescription && m_DropDescription.SetInsert(lpszText);
}

// Set DropDescription image type and text.
// This may be called from OnDragEnter() and OnDragOver() of control window classes.
//
// When not calling this from a control window class, it is called here when m_bUseDropDescription is true
//  to set the description text.
// NOTES: 
// - When calling this with nImageType == DROPIMAGE_INVALID, 
//    the description text strings are cleared. Used here from OnDragLeave().
// - When setting bCreate to true and the source provides a drag image but did not handle drop descriptions,
//    old and new style cursors may be shown together. See also the notes below in the code.
//
// Returns true when the drop description data object exists already or has been created.
bool COleDropTargetEx::SetDropDescription(DROPIMAGETYPE nImageType, LPCWSTR lpszText, bool bCreate)
{
	ASSERT(m_lpDataObject);

	bool bHasDescription = false;
	// Source provides a drag image, Vista or later with enabled themes,
	//  and actually dragging (m_lpDataObject is NULL when not dragging).
	if (m_bHasDragImage && m_bCanShowDescription && m_lpDataObject)
	{
		STGMEDIUM StgMedium;
		FORMATETC FormatEtc;
		if (CDragDropHelper::GetGlobalData(m_lpDataObject, CFSTR_DROPDESCRIPTION, FormatEtc, StgMedium))
		{
			bHasDescription = true;
			bool bUpdate = false;
			DROPDESCRIPTION *pDropDescription = static_cast<DROPDESCRIPTION*>(::GlobalLock(StgMedium.hGlobal));
			// Clear description
			if (DROPIMAGE_INVALID == nImageType)
				bUpdate = m_DropDescription.ClearDescription(pDropDescription);
			// There is no need to update the description text when the source has not enabled display of text.
			else if (m_bTextAllowed)
			{
				// If no text has been passed, use the stored text for the image type if present.
				if (NULL == lpszText)
					lpszText = m_DropDescription.GetText(nImageType, m_DropDescription.HasInsert());
				bUpdate = m_DropDescription.SetDescription(pDropDescription, lpszText);
			}
			if (pDropDescription->type != nImageType)
			{
				bUpdate = true;
				pDropDescription->type = nImageType;
			}
			::GlobalUnlock(StgMedium.hGlobal);
			if (bUpdate)
				bUpdate = SUCCEEDED(m_lpDataObject->SetData(&FormatEtc, &StgMedium, TRUE));
			if (!bUpdate)							// Nothing changed or setting data failed
				::ReleaseStgMedium(&StgMedium);
		}
		// DropDescription data object does not exist yet.
		// So create it when bCreate is true and image type is not invalid.
		// When m_bTextAllowed is true, we can be sure that the source supports drop descriptions.
		// Otherwise, there is no need to create the object when the image type
		//  corresponds to the drop effects.
		// This avoids creating the description object when the source does
		//  not handle them resulting in old and new style cursors shown together.
		// However, when using special image types (> DROPIMAGE_LINK), the object is
		//  always created which may result in showing old and new style cursors when the
		//  source does not support drop descriptions.
		if (!bHasDescription && bCreate &&
			DROPIMAGE_INVALID != nImageType &&
			(m_bTextAllowed || nImageType > DROPIMAGE_LINK))
//			(m_bTextAllowed || nImageType != static_cast<DROPIMAGETYPE>(dropEffect & ~DROPEFFECT_SCROLL)))
		{
			StgMedium.tymed = TYMED_HGLOBAL;
			StgMedium.hGlobal = ::GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, sizeof(DROPDESCRIPTION));
			StgMedium.pUnkForRelease = NULL;
			if (StgMedium.hGlobal)
			{
				DROPDESCRIPTION *pDropDescription = static_cast<DROPDESCRIPTION*>(::GlobalLock(StgMedium.hGlobal));
				pDropDescription->type = nImageType;
				if (m_bTextAllowed)
				{
					// If no text has been passed, use the stored text for the image type if present.
					if (NULL == lpszText)
						lpszText = m_DropDescription.GetText(nImageType, m_DropDescription.HasInsert());
					m_DropDescription.SetDescription(pDropDescription, lpszText);
				}
				::GlobalUnlock(StgMedium.hGlobal);
				bHasDescription = SUCCEEDED(m_lpDataObject->SetData(&FormatEtc, &StgMedium, TRUE));
				if (!bHasDescription)
					::GlobalFree(StgMedium.hGlobal);
			}
		}
	}
	// Indicate that this function has been called.
	// If it has been called by a window class that uses this class as member,
	//  it is not called again here for the current drag event.
	m_bDescriptionUpdated = true;
	return bHasDescription;
}

// Set the drop description data to the image type matching the drop effect.
// This may be called from OnDragEnter() and OnDragOver() of control window classes.
// When not calling this from a window class, it is called here when m_bUseDropDescription is true.
// This will use the default cursor type according to the drop effect.
// So there is no need to update the drop description when text is disabled.
// Returns true when the drop description data object exists already or has been created.
bool COleDropTargetEx::SetDropDescription(DROPEFFECT dwEffect)
{
	bool bHasDescription = false;
	if (m_bTextAllowed)								// Drag image present, Vista or later, 
	{												//  visual styles enabled, and text enabled
		// Filtering should have been already done by the window class when called from there.
		// When called from here (OnDragEnter, OnDragOver) it has been done.
		// This ensures that the correct drop description is shown.
		DROPIMAGETYPE nType = static_cast<DROPIMAGETYPE>(FilterDropEffect(dwEffect & ~DROPEFFECT_SCROLL));
		LPCWSTR lpszText = m_DropDescription.GetText(nType, m_DropDescription.HasInsert());
		if (lpszText)								// text is present (even if empty)
			bHasDescription = SetDropDescription(nType, lpszText, true);	
		else										// show default cursor and text
			bHasDescription = ClearDropDescription();
	}
	m_bDescriptionUpdated = true;					// don't call this again from OnDragEnter() or OnDragOver()
	return bHasDescription;
}

// Called by OLE when drag cursor first enters a registered window.
// Use this for:
// - Determination if target is able to drop data.
// - Optional highlight target.
// - Optional set drop description cursor type and text
DROPEFFECT COleDropTargetEx::OnDragEnter(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point)
{
	DROPEFFECT dwRet = DROPEFFECT_NONE;				// Default return value if not handled.
	ClearStateFlags();								// Clear drag session related flags.
	// Check if the source provides a drag image.
	m_bHasDragImage = (0 != CDragDropHelper::GetGlobalDataDWord(m_lpDataObject, _T("DragWindow")));
	// Check if the drag source has display of description text enabled.
	// If display of text is disabled, there is no need to change the description text here.
	if (m_bHasDragImage && m_bCanShowDescription)
		m_bTextAllowed = (DSH_ALLOWDROPDESCRIPTIONTEXT == CDragDropHelper::GetGlobalDataDWord(m_lpDataObject, _T("DragSourceHelperFlags")));
	// Cache preferred drop effect.
	m_nPreferredDropEffect = CDragDropHelper::GetGlobalDataDWord(m_lpDataObject, CFSTR_PREFERREDDROPEFFECT);

	if (m_pOnDragEnter)								// Callback function is provided
	{
		dwRet = m_pOnDragEnter(pWnd, pDataObject, dwKeyState, point);
	}
	else if (m_bUseMsg)								// Using messages
	{
		DragParameters Params = { 0 };
		Params.dwKeyState = dwKeyState;
		Params.point = point;
		dwRet = static_cast<DROPEFFECT>(pWnd->SendMessage(
			WM_APP_DRAG_ENTER, 
			reinterpret_cast<WPARAM>(pDataObject),
			reinterpret_cast<LPARAM>(&Params)));
	}
	else											// Default handling (CView support).
	{
		dwRet = COleDropTarget::OnDragEnter(pWnd, pDataObject, dwKeyState, point);
	}
#if 0
	// Example code: call handlers for supported classes
	if (pWnd->IsKindOf(RUNTIME_CLASS(CMyClass)))
	{
		CMyClass *pCast= static_cast<CMyClass*>(pWnd);
		ASSERT_VALID(pCast);
		dwRet = pCast->OnDragEnter(pDataObject, dwKeyState, point);
	}
#endif
	// COleDropTarget::XDropTarget::DragEnter() adjusts the effect if necessary
	//   after calling this function using _AfxFilterDropEffect().
	// But we need the correct effect here for the drop description and
	//  IDropTargetHelper::DragEnter().
	dwRet = FilterDropEffect(dwRet);
	// Set drop description text if available and not been already set by the window class.
	if (m_bUseDropDescription && !m_bDescriptionUpdated)
		SetDropDescription(dwRet);
	if (m_pDropTargetHelper)								// show drag image
		m_pDropTargetHelper->DragEnter(pWnd->GetSafeHwnd(), m_lpDataObject, &point, dwRet);
	m_bEntered = true;								// Entered the target window. Used by OnDragScroll().
	return dwRet;
}

// Called when drag cursor leaves a registered window.
// Use this for cleanup (e.g. remove highlighting).
// NOTE: 
//  This is not called when OnDrop or OnDropEx are implemented in a window class and has been called!
//  So it may be necessary to perform a similar clean up at the end of OnDrop or OnDropEx of the window class!
void COleDropTargetEx::OnDragLeave(CWnd* pWnd)
{
	if (m_pOnDragLeave)								// Use callback function
		m_pOnDragLeave(pWnd);
	else if (m_bUseMsg)								// Using messages
		pWnd->SendMessage(WM_APP_DRAG_LEAVE);
	else
		COleDropTarget::OnDragLeave(pWnd);
#if 0
	// Example code: call handlers for supported classes
	if (pWnd->IsKindOf(RUNTIME_CLASS(CMyClass)))
	{
		CMyClass *pCast= static_cast<CMyClass*>(pWnd);
		ASSERT_VALID(pCast);
		pCast->OnDragLeave();
	}
#endif
	// We should always set the image type of existing drop descriptions to DROPIMAGE_INVALID
	//  upon leaving even when the description hasn't been changed by this class.
	// This helps the drop source to detect changings of the description.
	ClearDropDescription();
	if (m_pDropTargetHelper)
		m_pDropTargetHelper->DragLeave();
	ClearStateFlags();								// Clear state flags. Next drag event may be from another source.
}

// Called when cursor is dragged over a registered window and
//  before OnDropEx/OnDrop when mouse button is released.
// Note: OnDragOver is like the WM_MOUSEMOVE message, it is called repeatedly
//   when the mouse moves inside the window.
// The mainly purpose of this function is to change the cursor according to
//  the key state.
// When dropping, the return value is passed to OnDropEx/OnDrop.
DROPEFFECT COleDropTargetEx::OnDragOver(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point)
{
	DROPEFFECT dwRet = DROPEFFECT_NONE;				// Default return value if not handled.
	m_bDescriptionUpdated = false;					// Set when description updated by target window
	if (m_pOnDragOver)
	{
		dwRet = m_pOnDragOver(pWnd, pDataObject, dwKeyState, point);
	}
	else if (m_bUseMsg)								// Using messages
	{
		DragParameters Params = { 0 };
		Params.dwKeyState = dwKeyState;
		Params.point = point;
		dwRet = static_cast<DROPEFFECT>(pWnd->SendMessage(
			WM_APP_DRAG_OVER,
			reinterpret_cast<WPARAM>(pDataObject),
			reinterpret_cast<LPARAM>(&Params)));
	}
	else											// Default handling (CView support).
	{
		dwRet = COleDropTarget::OnDragOver(pWnd, pDataObject, dwKeyState, point);
	}
#if 0
	// Example code: call handlers for supported classes
	if (pWnd->IsKindOf(RUNTIME_CLASS(CMyClass)))
	{
		CMyClass *pCast= static_cast<CMyClass*>(pWnd);
		ASSERT_VALID(pCast);
		dwRet = pCast->OnDragOver(pDataObject, dwKeyState, point);
	}
#endif
	// COleDropTarget::XDropTarget::DragEnter() adjusts the effect if necessary
	//   after calling this function using _AfxFilterDropEffect().
	// But we need the correct effect here for the drop description and
	//  IDropTargetHelper::DragEnter().
	dwRet = FilterDropEffect(dwRet);
	// Set drop description text if available and not been already set by the window class.
	if (m_bUseDropDescription && !m_bDescriptionUpdated)
		SetDropDescription(dwRet);
	if (m_pDropTargetHelper)
		m_pDropTargetHelper->DragOver(&point, dwRet);		// Show drag image
	return dwRet;
}

// Called when mouse button is released.
// Before this is called, OnDragOver() is called and the return value is passed in the dropDefault parameter.
// Note that this is always called even when OnDragOver() returned DROPEFFECT_NONE.
// So implementations should check dropDefault and do nothing if it is DROPEFFECT_NONE.
// They may optionally return -1 in this case to get OnDragLeave() called.
DROPEFFECT COleDropTargetEx::OnDropEx(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point)
{           
	DROPEFFECT dwRet = (DROPEFFECT)-1;				// Default return value if not handled.
	if (m_pOnDropEx)
	{
		dwRet = m_pOnDropEx(pWnd, pDataObject, dropDefault, dropList, point);
	}
	else if (m_bUseMsg)								// Using messages
	{												// Note that this handler must be always implemented!
		DragParameters Params = { 0 };				// If the handler does not handle the event, it must return -1.
		Params.dropEffect = dropDefault;
		Params.dropList = dropList;
		Params.point = point;
		dwRet = static_cast<DROPEFFECT>(pWnd->SendMessage(
			WM_APP_DROP_EX, 
			reinterpret_cast<WPARAM>(pDataObject),
			reinterpret_cast<LPARAM>(&Params)));
	}
	else											// Default handling (CView support).
	{
		dwRet = COleDropTarget::OnDropEx(pWnd, pDataObject, dropDefault, dropList, point);
	}
#if 0
	// Example code: call handlers for supported classes
	if (pWnd->IsKindOf(RUNTIME_CLASS(CMyClass)))
	{
		CMyClass *pCast= static_cast<CMyClass*>(pWnd);
		ASSERT_VALID(pCast);
		dwRet = pCast->OnDropEx(pDataObject, dropDefault, dropList, point);
	}
#endif
	if ((DROPEFFECT)-1 != dwRet)					// has been handled by window
		OnPostDrop(pDataObject, dwRet, point);
	// When dwRet == -1 and dropDefault != DROPEFFECT_NONE, OnDrop() is called now.
	// When dwRet == -1 and dropDefault == DROPEFFECT_NONE, OnDragLeave() is called now.
	return dwRet;
}

// Called when mouse button is released, OnDropEx() returns -1 and dropEffect != DROPEFFECT_NONE.
// dropEffect is the value returned by OnDragOver() called just before this.
// OnDragLeave() is not called. So it may be necessary to perform cleanup here.
// Return TRUE when data has been dropped.
// The drop effect passed to the DoDragDrop() call of the source is dropEffect 
//  when returning TRUE here or DROPEFFECT_NONE when returning FALSE.
BOOL COleDropTargetEx::OnDrop(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point)
{           
	BOOL bRet = FALSE;								// Default return value if not handled.
	if (m_pOnDrop)
	{
		bRet = m_pOnDrop(pWnd, pDataObject, dropEffect, point);
	}
	else if (m_bUseMsg)								// Using messages
	{
		DragParameters Params = { 0 };
		Params.dropEffect = dropEffect;
		Params.point = point;
		bRet = static_cast<BOOL>(pWnd->SendMessage(
			WM_APP_DROP, 
			reinterpret_cast<WPARAM>(pDataObject),
			reinterpret_cast<LPARAM>(&Params)));
	}
	else											// Default handling (CView support).
	{
		bRet = COleDropTarget::OnDrop(pWnd, pDataObject, dropEffect, point);
	}
#if 0
	// Example code: call handlers for supported classes
	if (pWnd->IsKindOf(RUNTIME_CLASS(CMyClass)))
	{
		CMyClass *pCast= static_cast<CMyClass*>(pWnd);
		ASSERT_VALID(pCast);
		dwRet = pCast->OnDrop(pDataObject, dropEffect, point);
	}
#endif
	OnPostDrop(pDataObject, bRet ? dropEffect : DROPEFFECT_NONE, point);
	return bRet;
}

// Internal function called when a drop operation has been performed.
void COleDropTargetEx::OnPostDrop(COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point)
{
	// CFSTR_PERFORMEDDROPEFFECT should be set by the window class!
	// When the source provides a CFSTR_PREFERREDDROPEFFECT object,
	//  we should set the performed drop effect if not done by the window class.
	if (DROPEFFECT_NONE != m_nPreferredDropEffect &&
		!pDataObject->IsDataAvailable(CDragDropHelper::RegisterFormat(CFSTR_PERFORMEDDROPEFFECT)))
	{
		CDragDropHelper::SetGlobalDataDWord(m_lpDataObject, CFSTR_PERFORMEDDROPEFFECT, dropEffect);
	}
	if (m_pDropTargetHelper)								// release drag image
		m_pDropTargetHelper->Drop(m_lpDataObject, &point, dropEffect);
	ClearStateFlags();								// Clear state flags. OnDragLeave() is not called.
}

// Optional auto scroll.
// This is called before OnDragEnter() and OnDragOver().
// Implemented functions should return:
// - DROPEFFECT_NONE:   
//    When not scrolling or not handling the event.
// - DROPEFFECT_SCROLL: 
//    When scrolling should be handled here according to the m_nScrollMode flags.
// - DROPEFFECT_SCROLL ored with DROPEFFECT_COPY, DROPEFFECT_MOVE, or DROPEFFECT_LINK. 
//    When scrolling is active.
// NOTE: 
//  When the value returned here has the DROPEFFECT_SCROLL bit set indicating
//   that scrolling is active, OnDragEnter() and OnDragOver() are not called afterwards.
DROPEFFECT COleDropTargetEx::OnDragScroll(CWnd* pWnd, DWORD dwKeyState, CPoint point)
{
	DROPEFFECT dwEffect = DROPEFFECT_NONE;
	// Use default scrolling (no OnDragScroll handler necessary)
	if (m_nScrollMode & AUTO_SCROLL_DEFAULT)
	{
		if (m_nScrollMode & AUTO_SCROLL_MASK)
			dwEffect = DoDefaultScroll(pWnd, dwKeyState, point);
	}
	else
	{
		if (m_pOnDragScroll)
		{
			dwEffect = m_pOnDragScroll(pWnd, dwKeyState, point);
		}
		else if (m_bUseMsg)							// Using messages
		{
			dwEffect = static_cast<DROPEFFECT>(pWnd->SendMessage(
				WM_APP_DRAG_SCROLL,
				dwKeyState,
				MAKELPARAM(point.x, point.y)));
		}
		else										// Default handling (CView support).
		{
			dwEffect = COleDropTarget::OnDragScroll(pWnd, dwKeyState, point);
		}
#if 0
		// Example code: call handlers for supported classes
		if (pWnd->IsKindOf(RUNTIME_CLASS(CMyClass)))
		{
			CMyClass *pCast= static_cast<CMyClass*>(pWnd);
			ASSERT_VALID(pCast);
			dwEffect = pCast->OnDragScroll(dwKeyState, point);
		}
#endif
		// When DROPEFFECT_SCROLL is returned, call our default handler.
		if (DROPEFFECT_SCROLL == dwEffect && (m_nScrollMode & AUTO_SCROLL_MASK))
			dwEffect = DoDefaultScroll(pWnd, dwKeyState, point);
	}

	// When the DROPEFFECT_SCROLL bit is set, OnDragEnter() or OnDragOver() are not called now.
	// Then the drag image must be updated here.
	if ((dwEffect & DROPEFFECT_SCROLL) && m_pDropTargetHelper)
	{
		// Adjust drop effect to show correct drop description cursor.
		// Our COleDropSourceEx class handles these correctly, but other
		//  sources may fail when the DROPEFFECT_SCROLL bit is set.
		dwEffect = FilterDropEffect(dwEffect & ~DROPEFFECT_SCROLL);
		if (m_bEntered)
			m_pDropTargetHelper->DragOver(&point, dwEffect);
		else
			m_pDropTargetHelper->DragEnter(pWnd->GetSafeHwnd(), m_lpDataObject, &point, dwEffect);
		dwEffect |= DROPEFFECT_SCROLL;
	}
	return dwEffect;
}

// Default auto scrolling.
// The second part of the code is based on the MFC COleDropTarget::OnDragScroll code for CView windows.
// Support for scroll bars is not included with COleDropTarget::OnDragScroll.
// NOTES:
// - COleDropTarget uses only the low word of m_nTimerID and checks this by comparing with 0xffff.
//   So we *MUST* do it the same way here when using m_nTimerID.
DROPEFFECT COleDropTargetEx::DoDefaultScroll(CWnd* pWnd, DWORD dwKeyState, CPoint point)
{
	ASSERT_VALID(pWnd);

	int nOnScrollBar = 0;
	int nVert = 0;
	int nHorz = 0;
	SCROLLBARINFO SbHorz, SbVert;

	// Get scroll bar info.
	// Also used with inset regions to check if scrolling in specific directions is possible.
	SbHorz.cbSize = sizeof(SbHorz);
	SbVert.cbSize = sizeof(SbVert);
	if (!pWnd->GetScrollBarInfo(OBJID_VSCROLL, &SbVert))
		SbVert.rgstate[0] = STATE_SYSTEM_INVISIBLE;
	if (!pWnd->GetScrollBarInfo(OBJID_HSCROLL, &SbHorz))
		SbHorz.rgstate[0] = STATE_SYSTEM_INVISIBLE;

	// Check if over a visible scroll bar.
	// Scroll if on an arrow or not over the thumb and it can be scrolled into the desired direction.
	// SCROLLBARINFO states for the scroll bar itself (index 0):
	// - STATE_SYSTEM_INVISIBLE: Scrollbar not visible
	// - STATE_SYSTEM_UNAVAIL:   Scrollbar grayed out (not enough content for scrolling)
	// - STATE_SYSTEM_NORMAL:    Can scroll in at least one direction
	// SCROLLBARINFO states for the thumb slider (indexes 2 and 4 for page up/right and down/left):
	//  Can't scroll in this direction if not STATE_SYSTEM_NORMAL. 
	CPoint ptScreen(point);
	pWnd->ClientToScreen(&ptScreen);
	if (IsScrollBarVisible(SbVert) && ::PtInRect(&SbVert.rcScrollBar, ptScreen))
	{
		nOnScrollBar = AUTO_SCROLL_VERT_BAR;
		// scroll bar must be enabled (not grayed out)
		if ((m_nScrollMode & AUTO_SCROLL_VERT_BAR) && IsScrollBarActive(SbVert))
		{
			int nPos = ptScreen.y - SbVert.rcScrollBar.top;
			if (CanScrollUpRight(SbVert) && nPos < SbVert.xyThumbTop)
				nVert = -1;
			else if (CanScrollDownLeft(SbVert) && nPos >= SbVert.xyThumbBottom)
				nVert = 1;
		}
	}
	else if (IsScrollBarVisible(SbHorz) && ::PtInRect(&SbHorz.rcScrollBar, ptScreen))
	{
		nOnScrollBar = AUTO_SCROLL_HORZ_BAR;
		if ((m_nScrollMode & AUTO_SCROLL_HORZ_BAR) && IsScrollBarActive(SbHorz))
		{
			int nPos = ptScreen.x - SbHorz.rcScrollBar.left;
			if (CanScrollDownLeft(SbHorz) && nPos < SbHorz.xyThumbTop)
				nHorz = -1;
			else if (CanScrollUpRight(SbHorz) && nPos >= SbHorz.xyThumbBottom)
				nHorz = 1;
		}
	}
	// Scroll if inside the inset region.
	else if (m_nScrollMode & AUTO_SCROLL_INSETS)
	{
		// Get client rectangle and non inset rectangle
		CRect rcClient;
		pWnd->GetClientRect(&rcClient);
		CRect rcNonInset = rcClient;
		// nScrollInset is a static COleDropTarget member variable.
		// Default inset hot zone width is 11 pixels.
		rcNonInset.InflateRect(-nScrollInset, -nScrollInset);
		if (rcClient.PtInRect(point) && !rcNonInset.PtInRect(point))
		{
			// Determine the scroll direction.
			// Check if scrolling can be performed according to scroll bar info:
			//  AUTO_SCROLL_INSET_BAR: True if scroll bar visible and enabled, and can scroll into direction.
			//  Otherwise: True if scroll bar invisible or can scroll into direction.
			if (m_nScrollMode & AUTO_SCROLL_INSET_HORZ)
			{
				if (point.x < rcNonInset.left)
				{
					if (CanScrollInsetDownLeft(SbHorz))
						nHorz = -1;
				}
				else if (point.x >= rcNonInset.right)
				{
					if (CanScrollInsetUpRight(SbHorz))
						nHorz = 1;
				}
			}
			if (m_nScrollMode & AUTO_SCROLL_INSET_VERT)
			{
				if (point.y < rcNonInset.top)
				{
					if (CanScrollInsetUpRight(SbVert))
						nVert = -1;
				}
				else if (point.y >= rcNonInset.bottom)
				{
					if (CanScrollInsetDownLeft(SbVert))
						nVert = 1;
				}
			}
		}
	}

	if (0 == nVert && 0 == nHorz)					// No scrolling
	{
		if (m_nTimerID != 0xffff)					// Was scrolling; so stop scrolling
		{
			COleDataObject dataObject;				// send fake OnDragEnter when changing from scroll to normal
			dataObject.Attach(m_lpDataObject, FALSE);
			OnDragEnter(pWnd, &dataObject, dwKeyState, point);
			m_nTimerID = 0xffff;					// no scrolling
		}
		return DROPEFFECT_NONE;
	}

	UINT nTimerID = MAKEWORD(nHorz, nVert);			// store scroll state as WORD 
	DWORD dwTick = ::GetTickCount();				// save tick count when timer ID changes
	if (nTimerID != m_nTimerID)						// start scrolling or direction has changed
	{
		m_dwLastTick = dwTick;
		m_nScrollDelay = nScrollDelay;				// scroll start delay time (default is 50 ms)
	}
	if (dwTick - m_dwLastTick > m_nScrollDelay)		// scroll if delay time expired
	{
		if (m_pScrollFunc)
			m_pScrollFunc(pWnd, nVert, nHorz);
		else if (m_bUseMsg)
			pWnd->SendMessage(WM_APP_DO_SCROLL, nVert, nHorz);
		else
		{
			// Try to access the scroll bar of the window.
			// NOTE: 
			//  This does not operate on scroll bars created when the WS_HSCROLL or WS_VSCROLL bits 
			//  are set during the creation of a window.
			if (nVert && IsScrollBarVisible(SbVert))
				SetScrollPos(pWnd, nVert, SB_VERT);
			if (nHorz && IsScrollBarVisible(SbHorz))
				SetScrollPos(pWnd, nHorz, SB_HORZ);
		}
		m_dwLastTick = dwTick;
		m_nScrollDelay = nScrollInterval;			// scrolling delay time (default is 50 ms)
	}
	if (0xffff == m_nTimerID && m_bEntered)			// scrolling started
	{
		OnDragLeave(pWnd);							// send fake OnDragLeave when changing from normal to scroll
	}
	m_nTimerID = nTimerID;							// save new scroll state
	// No dropping if over a scroll bar.
	// Otherwise return effect according to key state.
	return nOnScrollBar ? DROPEFFECT_SCROLL : GetDropEffect(dwKeyState) | DROPEFFECT_SCROLL;
}

// Try to access scroll bar of window.
// NOTE: 
//  This does not operate on scroll bars created when the WS_HSCROLL or WS_VSCROLL bits 
//  are set during the creation of a window.
void COleDropTargetEx::SetScrollPos(const CWnd* pWnd, int nStep, int nBar) const
{
	ASSERT_VALID(pWnd);
	ASSERT(nBar == SB_VERT || nBar == SB_HORZ);

	CWnd *pScrollWnd = pWnd->GetScrollBarCtrl(nBar);
	if (pScrollWnd && pScrollWnd->IsKindOf(RUNTIME_CLASS(CScrollBar)))
	{
		CScrollBar *pScrollCtrl = static_cast<CScrollBar*>(pScrollWnd);
		int nPos = pScrollCtrl->GetScrollPos() + nStep;
		pScrollCtrl->SetScrollPos(nPos);
	}
}

// Get drop effect from key state and allowed effects.
DROPEFFECT COleDropTargetEx::GetDropEffect(DWORD dwKeyState, DROPEFFECT dwDefault /*= DROPEFFECT_MOVE*/) const
{
	// Check for force link if linking is allowed.
	if ((m_nDropEffects & DROPEFFECT_LINK) && 
		(dwKeyState & (MK_CONTROL | MK_SHIFT)) == (MK_CONTROL | MK_SHIFT))
	{
		return DROPEFFECT_LINK;
	}
	// Check for force move if moving is allowed.
	if ((m_nDropEffects & DROPEFFECT_MOVE) && 
		((dwKeyState & MK_ALT) == MK_ALT || (dwKeyState & MK_SHIFT) == MK_SHIFT))
	{
		return DROPEFFECT_MOVE;
	}
	// Check for force copy (force move when dwDefault is copy)
	if ((dwKeyState & MK_CONTROL) == MK_CONTROL)
	{
		return FilterDropEffect((DROPEFFECT_COPY == dwDefault) ? DROPEFFECT_MOVE : DROPEFFECT_COPY);
	}
	// Default effect
	return FilterDropEffect((DROPEFFECT_NONE == dwDefault) ? DROPEFFECT_MOVE : dwDefault);
}

// Adjust effect according to allowed effects.
// COleDropTarget neither passes nor stores the allowed effects.
// The effect is adjusted if necessary after all handlers has been called
//  using _AfxFilterDropEffect().
// This may lead to wrong drop descriptions. So we have to store the allowed
//  effects in this class to be used here.
DROPEFFECT COleDropTargetEx::FilterDropEffect(DROPEFFECT dropEffect) const
{
	DROPEFFECT dropScroll = dropEffect & DROPEFFECT_SCROLL;
	dropEffect &= ~DROPEFFECT_SCROLL;
	if (dropEffect)
	{
#if 1
		// Mask out allowed effects
		dropEffect &= m_nDropEffects;
		// The source may have set preferred drop effect(s).
		if (dropEffect & m_nPreferredDropEffect)
			dropEffect &= m_nPreferredDropEffect;
		// Drop effect does not match the effects provided by the source.
		if (0 == dropEffect)
		{
			dropEffect = m_nPreferredDropEffect & m_nDropEffects;
			if (0 == dropEffect)
				dropEffect = m_nDropEffects;
		}
		// Choose a drop effect (multiple or no matching effects)
		if (dropEffect & DROPEFFECT_MOVE)
			dropEffect = DROPEFFECT_MOVE;
		else if (dropEffect & DROPEFFECT_COPY)
			dropEffect = DROPEFFECT_COPY;
		else if (dropEffect & DROPEFFECT_LINK)
			dropEffect = DROPEFFECT_LINK;
#else
		// Static function without declaration. See oledrop2.cpp.
		// NOTE: This does not handle effects with the DROPEFFECT_SCROLL bit set!
		AFX_STATIC DROPEFFECT AFXAPI _AfxFilterDropEffect(DROPEFFECT dropEffect, DROPEFFECT dwEffects);
		dropEffect = _AfxFilterDropEffect(dropEffect, m_nDropEffects);
#endif
	}
	return dropEffect | dropScroll;
}

// Functions to access drop data objects.
// These functions can be also used when pasting data from the clipboard.

// Get best matching available standard text format.
// With clipboard operations, standard text formats will be implicitly
//  converted if at least one format is available. But with drag & drop
//  operations this is not provided.
// Use this function to check if a standard text format is available and upon
//  success call GetString() passing the returned format to get the content
//  (as ANSI or Unicode string according to the project settings).
CLIPFORMAT COleDropTargetEx::GetBestTextFormat(COleDataObject* pDataObject) const
{ 
	ASSERT(pDataObject);

	CLIPFORMAT cfFormat = 0;
#ifdef _UNICODE
	if (pDataObject->IsDataAvailable(CF_UNICODETEXT))
		cfFormat = CF_UNICODETEXT;
	else if (pDataObject->IsDataAvailable(CF_TEXT))
		cfFormat = CF_TEXT;
#else
	if (pDataObject->IsDataAvailable(CF_TEXT))
		cfFormat = CF_TEXT;
	else if (pDataObject->IsDataAvailable(CF_UNICODETEXT))
		cfFormat = CF_UNICODETEXT;
#endif
	else if (pDataObject->IsDataAvailable(CF_OEMTEXT))
		cfFormat = CF_OEMTEXT;
	return cfFormat;
}

// Get the code page for ANSI and OEM text to Unicode conversions.
// If CF_LOCALE is not provided, the default system LCID is used.
// Must use the system LCID because data may be from other users.
UINT COleDropTargetEx::GetCodePage(COleDataObject* pDataObject, CLIPFORMAT cfFormat /*= CF_TEXT*/) const
{
	ASSERT(pDataObject);

	UINT nCP = CP_ACP;
	if (CF_UNICODETEXT == cfFormat)
		nCP = 1200;
	else
	{
		LCID nLCID = LOCALE_SYSTEM_DEFAULT;			// use sytem LCID if CF_LOCALE not present
		if (pDataObject->IsDataAvailable(CF_LOCALE))
		{
			HGLOBAL hGlobal = pDataObject->GetGlobalData(CF_LOCALE);
			if (hGlobal)
			{
				nLCID = *(static_cast<LCID *>(::GlobalLock(hGlobal)));
				::GlobalUnlock(hGlobal);
				::GlobalFree(hGlobal);
			}
		}
		nCP = CDragDropHelper::GetCodePage(nLCID, CF_OEMTEXT == cfFormat);
	}
	return nCP;
}

// Get copy of text data into allocated memory with size checking.
// The calling function must delete the returned memory.
// nCP must be 1200 for wide strings if format is not CF_UNICODETEXT.
// NOTE: Ensure that cfFormat is some kind of text format!
LPTSTR COleDropTargetEx::GetString(CLIPFORMAT cfFormat, COleDataObject* pDataObject, UINT nCP /*= 0*/) const
{
	ASSERT(cfFormat);
	ASSERT(pDataObject);

	size_t nDataSize = 0;
	LPTSTR lpszText = NULL;
	LPVOID lpData = GetData(pDataObject, cfFormat, nDataSize);
	if (lpData)
	{
		size_t nCopy = 0;
		if (1200 == nCP || CF_UNICODETEXT == cfFormat)
		{
			LPCWSTR lpszWide = static_cast<LPCWSTR>(lpData);
			nCopy = CDragDropHelper::GetWideLength(lpszWide, nDataSize);
#ifdef _UNICODE
			if (nCopy < nDataSize / sizeof(WCHAR))	// null terminated wide char string
			{
				lpszText = static_cast<LPTSTR>(lpData);
				nCopy = 0;
			}
#else
			// Convert to multi byte using the code page of the active thread.
			lpszText = CDragDropHelper::WideCharToMultiByte(lpszWide, static_cast<int>(nCopy));
			nCopy = 0;
#endif
		}
		else
		{
			LPCSTR lpszAnsi = static_cast<LPCSTR>(lpData);
			nCopy = CDragDropHelper::GetMultiByteLength(lpszAnsi, nDataSize);
			if (0 == nCP)
				nCP = GetCodePage(pDataObject, cfFormat);
#ifdef _UNICODE
			// Convert to wide string using the code page obtained from CF_LOCALE.
			lpszText = CDragDropHelper::MultiByteToWideChar(lpszAnsi, static_cast<int>(nCopy), nCP);
			nCopy = 0;
#else
			// If the code page is not identical to those of the current thread,
			//  text must be converted to Unicode and then to the current code page.
			// Examples are UTF-8 strings, CF_OEMTEXT format, and drag sources specifying a LCID
			//  that does not match the currently used one.
			if (nCP && !CDragDropHelper::CompareCodePages(nCP, CP_THREAD_ACP))
			{
				lpszText = CDragDropHelper::MultiByteToMultiByte(lpszAnsi, static_cast<int>(nCopy), nCP);
				nCopy = 0;
			}
			else if (nCopy < nDataSize)				// null terminated char string
			{
				lpszText = static_cast<LPTSTR>(lpData);
				nCopy = 0;
			}
#endif
		}
		if (nCopy)									// text has not been converted and is not null terminated
		{											// copy content from the buffer
			try
			{
				lpszText = new TCHAR[nCopy + 1];
				::CopyMemory(lpszText, lpData, nCopy * sizeof(TCHAR));
				lpszText[nCopy] = _T('\0');
			}
			catch (CMemoryException *pe)
			{
				pe->Delete();
			}
		}
		if (lpData != lpszText)						// data has been copied or converted
			delete [] lpData;
	}
	return lpszText;
}

// Get string from standard text formats.
// With clipboard operations, standard text formats will be implicitly
//  converted if at least one format is available. But with drag & drop
//  operations this is not provided.
LPTSTR COleDropTargetEx::GetText(COleDataObject* pDataObject) const
{
	CLIPFORMAT cfFormat = GetBestTextFormat(pDataObject);
	return cfFormat ? GetString(cfFormat, pDataObject) : NULL;
}

// Get length of text for standard text formats and optional count number of lines.
// Returns number of characters without terminating NULL byte.
// Note that the returned number of characters may be less than the number of bytes
//  with multi byte code pages.
size_t COleDropTargetEx::GetCharCount(COleDataObject* pDataObject, unsigned *pnLines) const
{
	CLIPFORMAT cfFormat = GetBestTextFormat(pDataObject);
	return cfFormat ? GetCharCount(cfFormat, pDataObject, pnLines) : 0;
}

// Get length of text for text formats and optional count number of lines.
// nCP must be 1200 for wide strings if format is not CF_UNICODETEXT.
// Returns number of characters without terminating NULL byte.
// Note that the returned number of characters may be less than the number of bytes
//  with multi byte code pages.
size_t COleDropTargetEx::GetCharCount(CLIPFORMAT cfFormat, COleDataObject* pDataObject, unsigned *pnLines, UINT nCP /*= 0*/) const
{
	ASSERT(cfFormat);
	ASSERT(pDataObject);

	unsigned nLines = 0;
	size_t nLen = 0;
	size_t nDataSize = 0;
	LPVOID lpData = GetData(pDataObject, cfFormat, nDataSize);
	if (lpData)
	{
		if (CF_UNICODETEXT == cfFormat || 1200 == nCP)
		{
			LPCWSTR lpszText = static_cast<LPCWSTR>(lpData);
			size_t nMaxLen = nDataSize / sizeof(WCHAR);
			while (lpszText[nLen] && nLen < nMaxLen)
			{
				if (L'\n' == lpszText[nLen])
					++nLines;
				++nLen;
			}
		}
		else
		{
			LPCSTR lpszText = static_cast<LPCSTR>(lpData);
			while (lpszText[nLen] && nLen < nDataSize)
			{
				if ('\n' == lpszText[nLen])
					++nLines;
				++nLen;
			}
			// Use MultiByteToWideChar() to get the number of characters for multi byte code pages.
			if (0 == nCP)
				nCP = GetCodePage(pDataObject, cfFormat);
			CPINFO CpInfo;
			::GetCPInfo(nCP, &CpInfo);
			if (CpInfo.MaxCharSize > 1)
				nLen = (size_t)::MultiByteToWideChar(nCP, 0, static_cast<LPCSTR>(lpData), static_cast<int>(nLen), NULL, 0);
		}
		if (nLen)
			++nLines;								// one more line than NL sequences
		delete [] lpData;
	}
	if (pnLines)
		*pnLines = nLines;
	return nLen;
}

// Get decoded HTML format content as allocated string.
// If bAll is true and the StartHTML and EndHTML keywords are valid, the complete HTML content is copied.
// Otherwise, the marked fragments are copied.
LPTSTR COleDropTargetEx::GetHtmlString(COleDataObject* pDataObject, bool bAll) const
{
	size_t nSize = 0;
	LPTSTR lpszHtml = NULL;
	LPCSTR lpData = GetHtmlData(pDataObject, nSize);
	if (lpData)
	{
		CHtmlFormat HtmlFormat;
		lpszHtml = HtmlFormat.GetHtml(lpData, static_cast<int>(nSize), bAll);
		delete [] lpData;
	}
	return lpszHtml;
}

// Get HTML format content as allocated UTF-8 string.
// If bAll is true and the StartHTML and EndHTML keywords are valid, the complete HTML content is copied.
// Otherwise, the marked fragments are copied.
LPSTR COleDropTargetEx::GetHtmlUtf8String(COleDataObject* pDataObject, bool bAll) const
{
	size_t nSize = 0;
	LPSTR lpszHtml = NULL;
	LPCSTR lpData = GetHtmlData(pDataObject, nSize);
	if (lpData)
	{
		CHtmlFormat HtmlFormat;
		lpszHtml = HtmlFormat.GetUtf8(lpData, static_cast<int>(nSize), bAll);
		delete [] lpData;
	}
	return lpszHtml;
}

// Get HTML format content as allocated UTF-8 string including the HTML format header.
LPCSTR COleDropTargetEx::GetHtmlData(COleDataObject* pDataObject, size_t& nSize) const
{
	CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(CF_HTML_STR);
	return pDataObject->IsDataAvailable(cfFormat) ? static_cast<LPCSTR>(GetData(pDataObject, cfFormat, nSize)) : NULL;
}

// Get single file name.
// If there are multiple file names, the name of the first file is returned.
CString COleDropTargetEx::GetSingleFileName(COleDataObject* pDataObject) const
{
	CString str;
	if (pDataObject->IsDataAvailable(CF_HDROP))
	{
		HGLOBAL hGlobal = pDataObject->GetGlobalData(CF_HDROP);
		if (hGlobal)
		{
			HDROP hDrop = static_cast<HDROP>(::GlobalLock(hGlobal));
			LPTSTR lpszName = str.GetBufferSetLength(MAX_PATH);
			unsigned nLen = ::DragQueryFile(hDrop, 0, lpszName, MAX_PATH);
			str.ReleaseBufferSetLength((int)nLen);
			::GlobalUnlock(hGlobal);
			::GlobalFree(hGlobal);
		}
	}
#if 0
	if (str.IsEmpty())
	{
		// Don't use CFSTR_FILENAMEA here. It is a short (8.3) name!
		CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(CFSTR_FILENAMEW);
		if (pDataObject->IsDataAvailable(cfFormat))
		{
			HGLOBAL hGlobal = pDataObject->GetGlobalData(cfFormat);
			if (hGlobal)
			{
				str = static_cast<LPCWSTR>(::GlobalLock(hGlobal));
				::GlobalUnlock(hGlobal);
				::GlobalFree(hGlobal);
			}
		}
	}
#endif
	return str;
}

// Get the number of files from CF_HDROP.
// CF_HDROP is present with various file list formats.
unsigned COleDropTargetEx::GetFileNameCount(COleDataObject* pDataObject) const
{
	unsigned nFiles = 0;
	if (pDataObject->IsDataAvailable(CF_HDROP))
	{
		HGLOBAL hGlobal = pDataObject->GetGlobalData(CF_HDROP);
		if (hGlobal)
		{
			HDROP hDrop = static_cast<HDROP>(::GlobalLock(hGlobal));
			nFiles = ::DragQueryFile(hDrop, UINT_MAX, NULL, 0);
			::GlobalUnlock(hGlobal);
			::GlobalFree(hGlobal);
		}
	}
	return nFiles;
}

// Get the number of files and length in characters for CR-LF seperated list of names from CF_HDROP.
size_t COleDropTargetEx::GetFileListAsTextLength(COleDataObject* pDataObject, unsigned *pnLines) const
{
	unsigned nLines = 0;
	size_t nLen = 0;
	if (pDataObject->IsDataAvailable(CF_HDROP))
	{
		HGLOBAL hGlobal = pDataObject->GetGlobalData(CF_HDROP);
		if (hGlobal)
		{
			HDROP hDrop = static_cast<HDROP>(::GlobalLock(hGlobal));
			nLen = GetFileListAsTextLength(hDrop, nLines);
			::GlobalUnlock(hGlobal);
			::GlobalFree(hGlobal);
		}
	}
	if (pnLines)
		*pnLines = nLines;
	return nLen;
}

// Get the number of files and length in characters for CR-LF seperated list of names from HDROP.
size_t COleDropTargetEx::GetFileListAsTextLength(HDROP hDrop, unsigned& nLines) const
{
	// Get number of file names
	nLines = ::DragQueryFile(hDrop, UINT_MAX, NULL, 0);
	size_t nLen = nLines ? 2 * (nLines - 1) : 0;	// CR-LF separated file names
	for (unsigned i = 0; i < nLines; i++)
		nLen += ::DragQueryFile(hDrop, i, NULL, 0);
	return nLen;
}

// Get list of file names from CF_HDROP as CR-LF separated string.
LPTSTR COleDropTargetEx::GetFileListAsText(COleDataObject* pDataObject) const
{
	LPTSTR lpszText = NULL;
	if (pDataObject->IsDataAvailable(CF_HDROP))
	{
		HGLOBAL hGlobal = pDataObject->GetGlobalData(CF_HDROP);
		if (hGlobal)
		{
			unsigned nFiles = 0;
			HDROP hDrop = static_cast<HDROP>(::GlobalLock(hGlobal));
			size_t nLen = GetFileListAsTextLength(hDrop, nFiles);
			if (nFiles)
			{
				try
				{
					lpszText = new TCHAR[nLen + 1];
					size_t nPos = 0;
					for (unsigned i = 0; i < nFiles; i++)
					{
						if (nPos)
						{
							lpszText[nPos++] = _T('\r');
							lpszText[nPos++] = _T('\n');
						}
						nPos += ::DragQueryFile(hDrop, i, lpszText + nPos, static_cast<UINT>(nLen + 1 - nPos));
					}
					ASSERT(nPos <= nLen);			// nPos == nLen expected
					lpszText[nPos] = _T('\0');
				}
				catch (CMemoryException *pe)
				{
					pe->Delete();
				}
			}
			::GlobalUnlock(hGlobal);
			::GlobalFree(hGlobal);
		}
	}
	return lpszText;
}

// Get DDB bitmap from CF_DIBV5, CF_DIB, CF_BITMAP, and various image formats.
HBITMAP COleDropTargetEx::GetBitmap(COleDataObject* pDataObject) const
{
	// First check if a DIB or DIBV5 bitmap is provided.
	HBITMAP hBitmap = GetBitmapFromDIB(pDataObject);
	// Check if a DDB bitmap is provided.
	if (NULL == hBitmap)
		hBitmap = GetBitmapFromDDB(pDataObject);
	// Check for various image types stored in IStream or file.
	// Supported formats: CF_TIFF and MIME formats image/tiff, image/png, image/gif, image/jpeg. 
	if (NULL == hBitmap)
		hBitmap = GetBitmapFromImage(pDataObject);
	return hBitmap;
}

// Get DDB bitmap from CF_BITMAP. 
HBITMAP COleDropTargetEx::GetBitmapFromDDB(COleDataObject* pDataObject) const
{
	ASSERT(pDataObject);

	HBITMAP hBitmap = NULL;
	if (pDataObject->IsDataAvailable(CF_BITMAP))
	{
		STGMEDIUM StgMedium;
		FORMATETC FormatEtc = { CF_BITMAP, NULL, DVASPECT_CONTENT, -1, TYMED_GDI };
		if (pDataObject->GetData(CF_BITMAP, &StgMedium, &FormatEtc))
		{
			// Return the handle from the medium when standard release is used (DeleteObject)
			if (NULL == StgMedium.pUnkForRelease)
				hBitmap = StgMedium.hBitmap;
			else
			{
				// Create a copy of the image.
				// OleDuplicateData creates a new GDI object
//				hBitmap = (HBITMAP)::CopyImage(StgMedium.hBitmap, IMAGE_BITMAP, 0, 0, 0);
				hBitmap = (HBITMAP)::OleDuplicateData(StgMedium.hBitmap, CF_BITMAP, 0);
				::ReleaseStgMedium(&StgMedium);
			}
		}
	}
	return hBitmap;
}

HPALETTE COleDropTargetEx::GetPalette(COleDataObject* pDataObject) const
{
	ASSERT(pDataObject);

	HPALETTE hPal = NULL;
	if (pDataObject->IsDataAvailable(CF_PALETTE))
	{
		STGMEDIUM StgMedium;
		FORMATETC FormatEtc = { CF_PALETTE, NULL, DVASPECT_CONTENT -1, TYMED_GDI };
		if (pDataObject->GetData(CF_PALETTE, &StgMedium, &FormatEtc))
		{
			if (NULL == StgMedium.pUnkForRelease)
				hPal = (HPALETTE)StgMedium.hBitmap;
			else
			{
				// OleDuplicateData creates a new GDI object
				hPal = (HPALETTE)::OleDuplicateData(StgMedium.hBitmap, CF_PALETTE, 0);
				::ReleaseStgMedium(&StgMedium);
			}
		}
	}
	return hPal;
}

// Get DDB bitmap from CF_DIBV5 or CF_DIB. 
HBITMAP COleDropTargetEx::GetBitmapFromDIB(COleDataObject* pDataObject) const
{
	ASSERT(pDataObject);

	HBITMAP hBitmap = NULL;
	BOOL bV5 = pDataObject->IsDataAvailable(CF_DIBV5);
	if (bV5 || pDataObject->IsDataAvailable(CF_DIB))
	{
		// Get DIB from HGLOBAL, FILE, or IStream.
		unsigned nSize;
		LPVOID lpData = GetData(pDataObject, bV5 ? CF_DIBV5 : CF_DIB, nSize); 
		if (lpData)
		{
			LPBITMAPINFOHEADER lpBIH = static_cast<LPBITMAPINFOHEADER>(lpData);
			unsigned nColorSize = lpBIH->biClrUsed * sizeof(RGBQUAD);
			if (lpBIH->biBitCount <= 8)
			{
				// If lpBI->bmiHeader.biClrUsed is non zero it specifies the number of colors.
				// If it is zero, all colors are used.
				if (0 == nColorSize)
					nColorSize = (lpBIH->biBitCount << 1) * sizeof(RGBQUAD);
			}
			// With BI_BITFIELDS compression, add 3 DWORDs.
			else if (BI_BITFIELDS == lpBIH->biCompression) 
				nColorSize = 3 * sizeof(DWORD);

			// Create a BITMAPINFO structure if the DIB has a V4 or V5 header.
			// Otherwise, we can use lpData which begins with a BITMAPINFO structure.
			LPBITMAPINFO lpBI = static_cast<LPBITMAPINFO>(lpData);
			// Use a stack based buffer with max. size here to avoid handling memory exceptions.
			LPBYTE lpBiBuf[sizeof(BITMAPINFO) + 255 * sizeof(RGBQUAD)];
			if (lpBIH->biSize != sizeof(BITMAPINFOHEADER))
			{
				lpBI = reinterpret_cast<LPBITMAPINFO>(lpBiBuf);
				// Copy the common content of the V4 or V5 header
				::CopyMemory(lpBI, lpBIH, sizeof(BITMAPINFO));
				// Adjust the size member
				lpBI->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				// Copy the color table
				if (nColorSize)
					::CopyMemory(lpBI->bmiColors, static_cast<LPBYTE>(lpData) + lpBIH->biSize, nColorSize);
			}

			HDC hScreen = ::GetDC(NULL);
			hBitmap = ::CreateDIBitmap(
				hScreen,							// Create screen compatible bitmap
				lpBIH,								// BITMAPINFOHEADER, BITMAPV4HEADER, or BITMAPV5HEADER
                CBM_INIT,							// Initialize bitmap with data from next parameters
                static_cast<LPBYTE>(lpData) + lpBIH->biSize + nColorSize, 
                lpBI,								// BITMAPINFO (BITMAPINFOHEADER followed by color table)
                DIB_RGB_COLORS);					// Color table uses RGB values
			::ReleaseDC(NULL, hScreen);
			delete [] lpData;
		}
	}
	return hBitmap;
}

// Get bitmap from various image types.
HBITMAP COleDropTargetEx::GetBitmapFromImage(COleDataObject* pDataObject) const
{
	ASSERT(pDataObject);

	HBITMAP hBitmap = NULL;
	if (pDataObject->IsDataAvailable(CF_TIFF))
		hBitmap = GetBitmapFromImage(pDataObject, CF_TIFF);
	else
	{
		// There is no standard for these clipboard formats.
		// The MIME types are often used with email and web applications.
		// "Windows Bitmap" is used by some applications.
		// Note that the bitmap format here is the BMP file format
		//  (which is a BITMAPFILEHEADER followed by a DIB).
		// Types that can be converted without loss of information should be listed first.
		LPCTSTR lpszFormats[] = { 
			_T("image/bmp"), _T("image/x-windows-bitmap"), _T("Windows Bitmap"), 
			_T("image/tiff"), _T("TIFF"), 
			_T("image/png"),  _T("PNG"), 
			_T("image/gif"),  _T("GIF"), 
			_T("image/jpeg"),  _T("JPEG"), 
			NULL 
		};
		int i = 0;
		while (NULL == hBitmap && lpszFormats[i])
		{
			CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(lpszFormats[i++]);
			if (pDataObject->IsDataAvailable(cfFormat))
				hBitmap = GetBitmapFromImage(pDataObject, cfFormat);
		}
	}
	return hBitmap;
}

// Get bitmap from various image types.
// This tries to load the data using the CImage class.
// Supported tymeds are HGLOBAL, FILE, ISTREAM, and MFPICT.
HBITMAP COleDropTargetEx::GetBitmapFromImage(COleDataObject* pDataObject, CLIPFORMAT cfFormat) const
{
	ASSERT(pDataObject);
	ASSERT(cfFormat);

	HBITMAP hBitmap = NULL;
	FORMATETC FormatEtc = { 0, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_FILE | TYMED_MFPICT};
	STGMEDIUM StgMedium;
	if (pDataObject->GetData(cfFormat, &StgMedium, &FormatEtc))
	{
		CImage Image;
		switch (StgMedium.tymed)
		{
		case TYMED_HGLOBAL :
//		case TYMED_MFPICT :
			{
				IStream *lpStream = NULL;
				if (SUCCEEDED(::CreateStreamOnHGlobal(StgMedium.hGlobal, FALSE, &lpStream)))
				{
					Image.Load(lpStream);
					lpStream->Release();
				}
			}
			break;
		case TYMED_FILE :
#ifdef _UNICODE
			Image.Load(StgMedium.lpszFileName);
#else
			{
				CString strFileName(StgMedium.lpszFileName);
				Image.Load(strFileName.GetString());
			}
#endif
			break;
		case TYMED_ISTREAM :
			Image.Load(StgMedium.pstm);
			break;
		}
		if (!Image.IsNull())
			hBitmap = Image.Detach();
		::ReleaseStgMedium(&StgMedium);
	}
	return hBitmap;
}

// Get a copy of the drag image bitmap.
// The global memory contains a SHDRAGIMAGE structure followed by the image bits.
// The SHDRAGIMAGE hbmpDragImage bitmap handle member is owned by OLE.
// So we can not use CopyImage() to create a copy of the bitmap because
//  GDI objects can not be shared between different processes.
HBITMAP COleDropTargetEx::GetDragImageBitmap(COleDataObject* pDataObject) const
{
	ASSERT(pDataObject);

	HBITMAP hBitmap = NULL;
	CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(_T("DragImageBits"));
	if (pDataObject->IsDataAvailable(cfFormat))
	{
		HGLOBAL hGlobal = pDataObject->GetGlobalData(cfFormat);
		if (hGlobal)
		{
			size_t nSize = ::GlobalSize(hGlobal);
			LPSHDRAGIMAGE pHdr = static_cast<LPSHDRAGIMAGE>(::GlobalLock(hGlobal));
			// Create a screen compatible bitmap using the size provided by the SHDRAGIMAGE structure.
			HDC hScreen = ::GetDC(NULL);
			if (hScreen)
				hBitmap = ::CreateCompatibleBitmap(hScreen, pHdr->sizeDragImage.cx, pHdr->sizeDragImage.cy);
			if (NULL != hBitmap)
			{
				BITMAP bm;
				VERIFY(sizeof(BITMAP) == ::GetObject(hBitmap, sizeof(bm), &bm));
				// If this is not set by GetObject().
				bm.bmWidthBytes = (((bm.bmWidth * bm.bmBitsPixel) + 31) / 32) * 4;

				// Provide buffer for BITMAPINFO and fill it.
				// When using dynamic allocation here, we must use a try / catch block
				//  to avoid memory leaks (global memory, DC, bitmap).
				// Because max. size is about 1 K, using a buffer on the stack is a better option.
				unsigned nColors = (bm.bmBitsPixel <= 8) ? (bm.bmBitsPixel << 1) : 0;
				LPBYTE pBiBuffer[sizeof(BITMAPINFOHEADER) + 255 * sizeof(RGBQUAD)];
				LPBITMAPINFO pBI = reinterpret_cast<LPBITMAPINFO>(pBiBuffer);
				// Fill the BITMAPINFOHEADER structure.
				::ZeroMemory(pBiBuffer, sizeof(BITMAPINFOHEADER));
				pBI->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				pBI->bmiHeader.biPlanes = 1;
				pBI->bmiHeader.biBitCount = bm.bmBitsPixel;
				pBI->bmiHeader.biCompression = BI_RGB;
				pBI->bmiHeader.biWidth = bm.bmWidth;
				pBI->bmiHeader.biHeight = bm.bmHeight;
				pBI->bmiHeader.biSizeImage = (DWORD)(bm.bmWidthBytes * bm.bmHeight);
				// Fill the color table using the system palette
				if (nColors)
					VERIFY(nColors == ::GetSystemPaletteEntries(hScreen, 0, nColors, reinterpret_cast<LPPALETTEENTRY>(pBI->bmiColors)));
				// Copy the bits.
				// With 64-bit apps there is an additional offset of 8 bytes.
				// See http://qa.social.msdn.microsoft.com/Forums/en-US/windowsuidevelopment/thread/1cbe3db8-70b0-4a19-88d5-334181d259ac
				LPBYTE lpBits = reinterpret_cast<LPBYTE>(pHdr) + sizeof(SHDRAGIMAGE);
				if (nSize == 8 + sizeof(SHDRAGIMAGE) + pBI->bmiHeader.biSizeImage)
					lpBits += 8;
				if (::SetDIBits(hScreen, hBitmap, 0, (UINT)bm.bmHeight, lpBits, pBI, DIB_RGB_COLORS) <= 0)
				{
					::DeleteObject(hBitmap);
					hBitmap = NULL;
				}
			}
			if (hScreen)
				::ReleaseDC(NULL, hScreen);
			::GlobalUnlock(hGlobal);
			::GlobalFree(hGlobal);
		}
	}
	return hBitmap;
}

// Copy data to allocated memory.
// Supported tymeds are HGLOBAL, FILE, ISTREAM, and MFPICT.
LPVOID COleDropTargetEx::GetData(COleDataObject* pDataObject, CLIPFORMAT cfFormat, size_t& nSize, LPFORMATETC lpFormatEtc /*= NULL*/) const
{
	LPBYTE lpData = NULL;
	// Returns CFile, CSharedFile, or COleStreamFile.
	// COleDataObject::GetFileData() catches exceptions.
	CFile *pFile = pDataObject->GetFileData(cfFormat, lpFormatEtc);
	nSize = 0;
	if (pFile)
	{
		nSize = static_cast<size_t>(pFile->GetLength());
		if (nSize)
		{
			try
			{
				lpData = new BYTE[nSize];
				LPBYTE lpRead = lpData;
				size_t nLeft = nSize;
				UINT nRead;
				do
				{
					size_t nBlock = min(nLeft, UINT_MAX);
					nRead = pFile->Read(lpRead, static_cast<UINT>(nBlock));
					lpRead += nRead;
					nLeft -= nRead;
				}
				while (nLeft && nRead);
			}
			catch (CMemoryException *pe)
			{
				nSize = 0;
				pe->Delete();
			}
		}
		// Close and delete the file
		delete pFile;
	}
	return lpData;
}

HGLOBAL COleDropTargetEx::GetDataAsGlobal(COleDataObject* pDataObject, CLIPFORMAT cfFormat, size_t& nSize, LPFORMATETC lpFormatEtc /*= NULL*/) const
{
	nSize = 0;
	HGLOBAL hGlobal = NULL;
	STGMEDIUM StgMedium;
	FORMATETC FormatEtc = { cfFormat, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL | TYMED_FILE | TYMED_ISTREAM };
	if (NULL == lpFormatEtc)
		lpFormatEtc = &FormatEtc;
	lpFormatEtc->tymed &= TYMED_HGLOBAL | TYMED_FILE | TYMED_ISTREAM;
	if (lpFormatEtc->tymed &&
		pDataObject->GetData(cfFormat, &StgMedium, lpFormatEtc))
	{
		bool bRelease = true;
		switch (StgMedium.tymed)
		{
		case TYMED_HGLOBAL :
			nSize = ::GlobalSize(StgMedium.hGlobal);
			if (NULL == StgMedium.pUnkForRelease)
			{
				hGlobal = StgMedium.hGlobal;
				bRelease = false;
			}
			else
				hGlobal = CDragDropHelper::CopyGlobalMemory(NULL, StgMedium.hGlobal, nSize);
			break;
		case TYMED_FILE :
			{
				HANDLE hFile = ::CreateFile(OLE2T(StgMedium.lpszFileName), GENERIC_READ, 0, NULL, OPEN_EXISTING, 0, NULL);
				if (INVALID_HANDLE_VALUE != hFile)
				{
					BY_HANDLE_FILE_INFORMATION fi;
					if (::GetFileInformationByHandle(hFile, &fi))
					{
						hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, fi.nFileSizeLow);
						if (hGlobal)
						{
							DWORD dwRead;
							BOOL bRet = ::ReadFile(hFile, ::GlobalLock(hGlobal), fi.nFileSizeLow, &dwRead, NULL);
							::GlobalUnlock(hGlobal);
							if (bRet)
								nSize = fi.nFileSizeLow;
							else
							{
								::GlobalFree(hGlobal);
								hGlobal = NULL;
							}
						}
					}
					::CloseHandle(hFile);
				}
			}
			break;
		case TYMED_ISTREAM :
			{
				STATSTG st;
				if (SUCCEEDED(StgMedium.pstm->Stat(&st, STATFLAG_NONAME)))
				{
					LARGE_INTEGER llNull;
					llNull.QuadPart = 0;
					if (SUCCEEDED(StgMedium.pstm->Seek(llNull, STREAM_SEEK_SET, NULL)))
					{
						hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, st.cbSize.LowPart);
						if (hGlobal)
						{
							ULONG nRead;
							bool bRet = SUCCEEDED(StgMedium.pstm->Read(::GlobalLock(hGlobal), st.cbSize.LowPart, &nRead));
							::GlobalUnlock(hGlobal);
							if (bRet)
								nSize = st.cbSize.LowPart;
							else
							{
								::GlobalFree(hGlobal);
								hGlobal = NULL;
							}
						}
					}
				}
			}
			break;
		}
		if (bRelease)
			::ReleaseStgMedium(&StgMedium);
	}
	return hGlobal;
}

// COleDropTarget neither passes nor stores the allowed effects.
// The effect is adjusted if necessary after all handlers has been called.
// But with drop descriptions, we need the allowed effects to perform 
// the adjustment before setting the drop description.
// So store the effects here in a member variable.

BEGIN_INTERFACE_MAP(COleDropTargetEx, COleDropTarget)
	INTERFACE_PART(COleDropTargetEx, IID_IDropTarget, DropTargetEx)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) COleDropTargetEx::XDropTargetEx::AddRef()
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) COleDropTargetEx::XDropTargetEx::Release()
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->ExternalRelease();
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return (HRESULT)pThis->ExternalQueryInterface(&iid, ppvObj);
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::DragEnter(THIS_ LPDATAOBJECT lpDataObject,
	DWORD dwKeyState, POINTL pt, LPDWORD pdwEffect)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	// Store allowed effects.
	pThis->m_nDropEffects = *pdwEffect;
	return pThis->m_xDropTarget.DragEnter(lpDataObject, dwKeyState, pt, pdwEffect);
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::DragOver(THIS_ DWORD dwKeyState,
	POINTL pt, LPDWORD pdwEffect)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	// Store allowed effects.
	pThis->m_nDropEffects = *pdwEffect;
	return pThis->m_xDropTarget.DragOver(dwKeyState, pt, pdwEffect);
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::DragLeave(THIS)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->m_xDropTarget.DragLeave();
}

STDMETHODIMP COleDropTargetEx::XDropTargetEx::Drop(THIS_ LPDATAOBJECT lpDataObject,
	DWORD dwKeyState, POINTL pt, LPDWORD pdwEffect)
{
	METHOD_PROLOGUE(COleDropTargetEx, DropTargetEx)
	return pThis->m_xDropTarget.Drop(lpDataObject, dwKeyState, pt, pdwEffect);
}

#ifdef _DEBUG
void COleDropTargetEx::Dump(CDumpContext& dc) const
{
	COleDropTarget::Dump(dc);

	dc << "m_bUseMsg = " << m_bUseMsg;
	dc << "\nm_bCanShowDescription = " << m_bCanShowDescription;
	dc << "\nm_bUseDropDescription = " << m_bUseDropDescription;
	dc << "\nm_bDescriptionUpdated = " << m_bDescriptionUpdated;
	dc << "\nm_bHasDragImage = " << m_bHasDragImage;
	dc << "\nm_bTextAllowed = " << m_bTextAllowed;
	dc << "\nm_bEntered = " << m_bEntered;
	dc << "\nm_nScrollMode = " << m_nScrollMode;
	dc << "\nm_nPreferredDropEffect = " << m_nPreferredDropEffect;
	dc << "\nm_nDropEffects = " << m_nDropEffects;

	dc << "\n";
}
#endif
