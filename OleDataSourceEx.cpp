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

// A COleDataSource derived class to copy data to the clipboard and using Drag&Drop.
// Supports:
// - Drag images.
// - New style (Vista) drag cursors with optional text (DropDescription).
// - Data preparation for plain text, CSV, HTML, and RTF.
// - Data preparation for DDB bitmaps (CF_BITMAP) and DIB bitmaps (CF_DIB, CF_DIBV5).
//
//
// NOTES: 
//  - Must call AfxOleInit() in InitInstance().
//  - Include afxole.h (best place would be stdafx.h)
//  - COleDataSource/COleDataSourceEx objects must be allocated from the heap using the new operator
//    when using clipboard functions like SetClipboard() (the object is destroyed automatically).
//  - Global memory objects must be allocated using GMEM_MOVEABLE.
//  - Drag & Drop does not work when drop target application is executed by another user.
//    Happens with Vista and later and enabled UAC.
//    Src = administrator, dst = user: 
//     Cursor indicates move/copy and drop operation returned, but nothing dropped.
//    Src = user, dst = administrator: 
//     No dropping possible (indicated by cursor).
//  - Code page / locale conversions are performed using the locale of the current thread.
//
// Should I use rendering mode?
//  Rendering mode should not be used for memory based clipboard operations.
//  With rendering, data will be requested when the copy operation is performed.
//  If the data has been changed meanwhile or the source window is no longer present,
//   changed data or even nothing would be copied.
//  With Drag & Drop operations rendering mode would avoid data preparation and
//   allocation if the operation is cancelled (nothing dropped) and consuming memory
//   when prividing data in multiple formats.
//  However, modern systems have much more RAM and CPU power than the systems
//   at the time when the clipboard and Drag & Drop support was introduced.
//  So there should be no need to use rendering mode.
//  Rendering may be used when providing file data that did not change. See also below.
//
// Should I use file data?
//  The historically reason to provide data via files is limited memory.
//  But modern systems have usually enough RAM to provide even large amounts of data
//  using memory.
//  Windows 9x: Global memory is limited to four mega byte blocks.
//
// Other helpful sources:
//
// http://www.codeproject.com/Articles/840/How-to-Implement-Drag-and-Drop-Between-Your-Progra
//
// MSDN: Shell Clipboard Formats
//  http://msdn.microsoft.com/en-us/library/windows/desktop/bb776902%28v=vs.85%29.aspx
// MSDN: Shell Data Object
//  http://msdn.microsoft.com/en-us/library/windows/desktop/bb776903%28v=vs.85%29.aspx

#include "StdAfx.h"
#include "OleDataSourceEx.h"

#ifdef OBJ_SUPPORT
#include <afxadv.h>									// CSharedFile class
#include ".\oledatasourceex.h"
#endif

// Link with uxtheme.lib. So there is no need to set it in the project settings.
#pragma comment(lib, "uxtheme")
#if (WINVER < 0x0501)
#pragma message("NOTE: Add uxtheme.dll to the list of delay loaded DLLs in the project settings.")
#endif

COleDataSourceEx::COleDataSourceEx(CWnd *pWnd /*= NULL*/, OnRenderDataFunc pOnRenderData /*= NULL*/)
{
	m_nDragResult = 0;								// COleDropSourceEx result flags.
	m_bSetDescriptionText = false;					// Set when drop description text string(s) has been specified
	m_pWnd = pWnd;									// Source window (passed to render callback function)
	m_pOnRenderData = pOnRenderData;				// Render callback function
	m_pDragSourceHelper = NULL;						// Drag source helper
	m_pDragSourceHelper2 = NULL;
	// Create drag source helper to show drag images.
#if defined(IID_PPV_ARGS) // The IID_PPV_ARGS macro has been introduced with Visual Studio 2005
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pDragSourceHelper));
#else
	::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER,
		IID_IDragSourceHelper, reinterpret_cast<LPVOID*>(&m_pDragSourceHelper));
#endif
	if (m_pDragSourceHelper)
	{
		// With Vista, IDragSourceHelper2 has been inherited from by IDragSourceHelper.
#if defined(IID_PPV_ARGS) && defined(__IDragSourceHelper2_INTERFACE_DEFINED__)
		// Can't use the IID_PPV_ARGS macro when building for pre Vista because the GUID for IDragSourceHelper2 is not defined.
		m_pDragSourceHelper->QueryInterface(IID_PPV_ARGS(&m_pDragSourceHelper2));
#else
		m_pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&m_pDragSourceHelper2));
#endif
		if (NULL != m_pDragSourceHelper2) 
		{
			// No need to have two instances of the source helper in memory.
			m_pDragSourceHelper->Release();
			m_pDragSourceHelper = static_cast<IDragSourceHelper*>(m_pDragSourceHelper2);
		}
	}
}

COleDataSourceEx::~COleDataSourceEx()
{
	if (NULL != m_pDragSourceHelper)
		m_pDragSourceHelper->Release();
}

// Perform drag & drop with support for DropDescription.
// To support drop descriptions, we must use our COleDropSource derived class
//  that handles drop descriptions within the GiveFeedback() function.
DROPEFFECT COleDataSourceEx::DoDragDropEx(DROPEFFECT dwEffect, LPCRECT lpRectStartDrag /*= NULL*/)
{
	// When visual styles are disabled, new cursors are not shown.
	// Then we must show the old style cursors.
	// NOTE: When pre XP Windows versions must be supported (WINVER < 0x0501),
	//  delay loading for uxtheme.dll must be specified in the project settings.
	bool bUseDescription = (m_pDragSourceHelper2 != NULL) && ::IsAppThemed();
	// When drop descriptions can be used and description text strings has been defined,
	//  create a DropDescription data object with image type DROPIMAGE_INVALID.
	if (bUseDescription && m_bSetDescriptionText)
	{												
		CLIPFORMAT cfDS = CDragDropHelper::RegisterFormat(CFSTR_DROPDESCRIPTION); 
		if (cfDS)
		{
			HGLOBAL hGlobal = ::GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, sizeof(DROPDESCRIPTION));
			if (hGlobal)
			{
				DROPDESCRIPTION *pDropDescription = static_cast<DROPDESCRIPTION*>(::GlobalLock(hGlobal));
				pDropDescription->type = DROPIMAGE_INVALID;
				::GlobalUnlock(hGlobal);
				CacheGlobalData(cfDS, hGlobal);
			}
		}
	}
	// Use our COleDropSource derived class here to support drop descriptions.
	COleDropSourceEx dropSource;
	if (bUseDescription)
	{
		dropSource.m_pIDataObj = static_cast<LPDATAOBJECT>(GetInterface(&IID_IDataObject));
		if (m_bSetDescriptionText)
			dropSource.m_pDropDescription = &m_DropDescription;
	}
	// Perform the drag operation.
	// This will block until the operation has finished.
	dwEffect = DoDragDrop(dwEffect, lpRectStartDrag, static_cast<COleDropSource*>(&dropSource));
	m_nDragResult = dropSource.m_nResult;			// Get the drag status
	if (dwEffect & ~DROPEFFECT_SCROLL)
		m_nDragResult |= DRAG_RES_DROPPED;			// Indicate that data has been successfully dropped
	return dwEffect;
}

// With Vista, IDragSourceHelper2 has been inherited from by IDragSourceHelper
//  providing the SetFlags() method.
// Call this after creation of this object to enable the display of description text with drag images.
// It must be called before SetDragImageWindow() or one of the SetDragImage() and InitDragImage() functions.
// When this succeeds, the global DWORD data object "DragSourceHelperFlags" is created.
bool COleDataSourceEx::AllowDropDescriptionText()
{
	return m_pDragSourceHelper2 ? SUCCEEDED(m_pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT)) : false;
}

// Pass drop description text.
// This text is used when target windows show drag images but did not supply their own description text.
// When no text is specified here and the target window did not set the text,
//  the system default text is used (as used by the Windows Explorer).
// nType:      Image type for which the text should be stored
// lpszText:   Text without using the insert string.
// lpszText1:  Text when using the insert string.
// lpszInsert: Insert text. Must be passed only once.
// Return:     True if text has been stored and can be displayed (Vista and later).
// NOTES:
//  - Because the %1 placeholder is used to insert the szInsert text, a single percent character
//    inside the strings must be esacped by another one.
//  - Don't forget to call AllowDropDescriptionText().
bool COleDataSourceEx::SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert /*= NULL*/)
{
	bool bRet = false;
	if (m_pDragSourceHelper2)
	{
		bRet = m_DropDescription.SetText(nType, lpszText, lpszText1);
		if (bRet && lpszInsert)
			m_DropDescription.SetInsert(lpszInsert);
		m_bSetDescriptionText |= bRet;
	}
	return bRet;
}

//
// Drag image functions
//

// Initialize the drag-image manager for a control with a window.
// The window must handle the registered DI_GETDRAGIMAGE message and fill the passed SHDRAGIMAGE structure.
// Built-in with CListCtrl and CTreeViewCtrl.
//
// Passing NULL for hWnd is allowed. Then a generic bitmap is created by the system.
// The generic bitmap is also used when the control window does not handle the DI_GETDRAGIMAGE message.
// In this case the boolean data object "UsingDefaultDragImage" is created and set to true.
// When a "Shell IDList Array" object exists, this is then used to determine the drag image
//  (like when dragging files from the Explorer).
bool COleDataSourceEx::SetDragImageWindow(HWND hWnd, POINT* pPoint)
{
	HRESULT hRes = E_NOINTERFACE;
	if (m_pDragSourceHelper)
	{
		POINT pt = { 0, 0 };
		if (NULL == pPoint)
			pPoint = &pt;
		hRes = m_pDragSourceHelper->InitializeFromWindow(hWnd, pPoint, 
			static_cast<LPDATAOBJECT>(GetInterface(&IID_IDataObject)));
#ifdef _DEBUG
		if (FAILED(hRes))
			TRACE1("COleDataSourceEx::SetDragImageWindow: InitializeFromWindow failed with code %#X\n", hRes);
#endif
	}
	return SUCCEEDED(hRes);
}

// Set drag image to a provided bitmap.
// pPoint: Cursor position relative to top left corner of image. 
//         If NULL, the position is set to the center of the image.
//         If a x or y position is negative, it is set to the center.
// clr:    Transparent color for the drag image (e.g. the background color of the bitmap).
//         With CLR_INVALID (-1), ::GetSystemColor(COLOR_WINDOW) is used.
//         To have non-transparent images, use a color that is not used by the bitmap.
//         With 24/32-bit images, Magenta (RGB(255, 0, 255)) is a common transparent color.
//         Note that pixels with this color are converted to black.
// NOTE: Ownership of hBitmap is passed to the DragHelper.
// The image can be accessed using the "DragImageBits" data object.
// See COleDropTargetEx::GetDragImageBitmap() to get a copy of the bitmap.
bool COleDataSourceEx::SetDragImage(HBITMAP hBitmap, const CPoint* pPoint, COLORREF clr)
{
	ASSERT(hBitmap);

	HRESULT hRes = E_NOINTERFACE;
	if (hBitmap && m_pDragSourceHelper)
	{
		BITMAP bm;
		SHDRAGIMAGE di;
		VERIFY(::GetObject(hBitmap, sizeof(bm), &bm));
		di.sizeDragImage.cx = bm.bmWidth;
		di.sizeDragImage.cy = bm.bmHeight;
		// If a position is negative, it is set to the center.
		// If a position is outside the image, it is set to the right / bottom.
		if (pPoint)
		{
			di.ptOffset = *pPoint;
			if (di.ptOffset.x < 0)
				di.ptOffset.x = di.sizeDragImage.cx / 2;
			else if (di.ptOffset.x > di.sizeDragImage.cx)
				di.ptOffset.x = di.sizeDragImage.cx;
			if (di.ptOffset.y < 0)
				di.ptOffset.y = di.sizeDragImage.cy / 2;
			else if (di.ptOffset.y > di.sizeDragImage.cy)
				di.ptOffset.y = di.sizeDragImage.cy;
		}
		// Center the cursor if pPoint is NULL. 
		else
		{
			di.ptOffset.x = di.sizeDragImage.cx / 2;
			di.ptOffset.y = di.sizeDragImage.cy / 2;
		}
		di.hbmpDragImage = hBitmap;
		di.crColorKey = (CLR_INVALID == clr) ? ::GetSysColor(COLOR_WINDOW) : clr;
		hRes = m_pDragSourceHelper->InitializeFromBitmap(&di,
			static_cast<LPDATAOBJECT>(GetInterface(&IID_IDataObject)));
#ifdef _DEBUG
		if (FAILED(hRes))
			TRACE1("COleDataSourceEx::SetDragImage: InitializeFromBitmap failed with code %#X\n", hRes);
#endif
	}
	if (FAILED(hRes) && hBitmap)
		::DeleteObject(hBitmap);					// delete image to avoid a memory leak
	return SUCCEEDED(hRes);
}

// Set drag image using bitmap loaded from resources.
bool COleDataSourceEx::InitDragImage(int nResBM, const CPoint* pPoint, COLORREF clr)
{
	bool bRet = false;
	if (m_pDragSourceHelper)
	{
		HBITMAP hBitmap = ::LoadBitmap(AfxGetResourceHandle(), MAKEINTRESOURCE(nResBM));
		ASSERT(NULL != hBitmap);					// resource not found
		if (clr && hBitmap)							// if transparent color is not black
			hBitmap = ReplaceBlack(hBitmap);		// replace black pixels by RGB(1, 1, 1)
		if (hBitmap)
			bRet = SetDragImage(hBitmap, pPoint, clr);
	}
	return bRet;
}

// Set drag image using file type icon.
// lpszFileExt: Valid file name extension (file must not exist, just pass something like ".txt").
// nScale:      Optional scaling (> 0 = percentage; < 0 = limit to max. width/height).
// pPoint:      Cursor position relative to left top corner of image. Centered if pPoint is NULL.
// nFlags:      Optional SHGetFileInfo flags
//  SHGFI_LARGEICON:     Retrieve the large icon (default).
//  SHGFI_LINKOVERLAY:   Add link symbol (not useful here with drag images). 
//  SHGFI_OPENICON:      Retrieve open icon (may be same as normal one).
//  SHGFI_SHELLICONSIZE: Get shell sized icon (may be same size as normal one).
//  SHGFI_SMALLICON:     Retrieve the small icon.
//  SHGFI_SELECTED:      Blend icon with system highlight color.
bool COleDataSourceEx::InitDragImage(LPCTSTR lpszExt, int nScale /*= 0*/, const CPoint* pPoint /*= NULL*/, UINT nFlags /*= 0*/)
{
	ASSERT(lpszExt && _tcschr(lpszExt, _T('.')));

	bool bRet = false;
	if (m_pDragSourceHelper)
	{
		SHFILEINFO fi;
		// Mask out flags that can't be used with SHGFI_ICON and SHGFI_USEFILEATTRIBUTES.
		nFlags &= ~(SHGFI_EXETYPE | SHGFI_ATTR_SPECIFIED | SHGFI_PIDL | SHGFI_ATTRIBUTES);
		// SHGFI_ICON: Get icon handle (must be destroyed).
		// SHGFI_USEFILEATTRIBUTES: Get file information by passing the extension.
		nFlags |= SHGFI_USEFILEATTRIBUTES | SHGFI_ICON;
#if 0
		if (nFlags & SHGFI_SYSICONINDEX)
		{
			nFlags &= ~SHGFI_ICON;
			HIMAGELIST himl = (HIMAGELIST)::SHGetFileInfo(lpszExt, FILE_ATTRIBUTE_NORMAL, &fi, sizeof(fi), nFlags);
			if (himl)
			{
				int cx, cy;
				ImageList_GetIconSize(himl, &cx, &cy);
				HDC hScreen = ::GetDC(NULL);
				HBITMAP hBitmap = ::CreateCompatibleBitmap(hScreen, cx, cy);
				HDC hDC = ::CreateCompatibleDC(hScreen);
				HBITMAP hOldBitmap = (HBITMAP)::SelectObject(hDC, hBitmap);
				ImageList_Draw(himl, fi.iIcon, hDC, 0, 0, ILD_NORMAL);
				VERIFY(hBitmap == (HBITMAP)::SelectObject(hDC, hOldBitmap));
				VERIFY(::DeleteDC(hDC));
				VERIFY(::ReleaseDC(NULL, hScreen));
				bRet = SetDragImage(hBitmap, pPoint, ImageList_GetBkColor(himl));
			}
		}
		else if (::SHGetFileInfo(lpszExt, FILE_ATTRIBUTE_NORMAL, &fi, sizeof(fi), nFlags))
#else
		if (::SHGetFileInfo(lpszExt, FILE_ATTRIBUTE_NORMAL, &fi, sizeof(fi), nFlags))
#endif
		{
			// Get icon info
			ICONINFO ii;
			VERIFY(::GetIconInfo(fi.hIcon, &ii));

			// Get the size of the icon from it's bitmap.
			// With a black & white icon, ii.hbmColor is NULL and
			//  the XOR mask is contained in the lower half of ii.hbmMask.
			BITMAP bm;
			VERIFY(::GetObject(ii.hbmColor ? ii.hbmColor : ii.hbmMask, sizeof(bm), &bm));
			CSize sz(bm.bmWidth, bm.bmHeight);
			if (NULL == ii.hbmColor)
				sz.cy /= 2;

			// Create a screen DC with the actual color depth
			HDC hScreen = ::GetDC(NULL);
			nScale = ScaleSize(nScale, sz, hScreen);
			// Create a compatible bitmap using the screen DC and select it into a memory DC.
			HBITMAP hBitmap = ::CreateCompatibleBitmap(hScreen, sz.cx, sz.cy);
			HDC hDC = ::CreateCompatibleDC(hScreen);
			HBITMAP hOldBitmap = (HBITMAP)::SelectObject(hDC, hBitmap);
			// Draw the icon.
			VERIFY(::DrawIconEx(hDC, 0, 0, fi.hIcon, sz.cx, sz.cy, 0, NULL, DI_NORMAL));
			// Cleanup
			VERIFY(hBitmap == (HBITMAP)::SelectObject(hDC, hOldBitmap));
			VERIFY(::DeleteDC(hDC));
			VERIFY(::ReleaseDC(NULL, hScreen));
			VERIFY(::DestroyIcon(fi.hIcon));
			bRet = SetDragImage(hBitmap, pPoint, 0);
		}
	}
	return bRet;
}

// Create drag image from text.
// Text is drawn using the same size and font as on the control.
// pWnd:          Control window.
// lpszText:      Text string.
// pPoint:        Cursor position relative to left top corner of image. Centered if pPoint is NULL.
// bHiLightColor: If true, use system highlight (selection) colors (no transparency).
//                If false, use text color of pWnd with transparent background.
bool COleDataSourceEx::InitDragImage(CWnd* pWnd, LPCTSTR lpszText, const CPoint *pPoint /*= NULL*/, bool bHiLightColor /*= false*/)
{
	ASSERT_VALID(pWnd);
	ASSERT(lpszText);

	bool bRet = false;
	if (pWnd && m_pDragSourceHelper && lpszText && *lpszText)
	{
		CDC memDC;
		CClientDC dc(pWnd);
		VERIFY(memDC.CreateCompatibleDC(&dc));

		// Create a copy of the font used by the control window.
		LOGFONT lf;
		CFont *pFont = pWnd->GetFont();
		VERIFY(pFont->GetLogFont(&lf));
		// Antialiasing should be disabled.
		// Otherwise, artifacts could occur at the edges of the drag image, 
		//  between the text color and the color key.
		lf.lfQuality = NONANTIALIASED_QUALITY;
#if 0
		if (nScale < 0)								// Limit font size to a maxmimum.
		{
			int nMaxHeight = ::MulDiv(nScale, memDC->GetDeviceCaps(LOGPIXELSY), 72);
			if (lf.lfHeight < nMaxHeight)			// both values are negative!
				lf.lfHeight = nMaxHeight;
		}
		else if (nScale > 0 && nScale != 100)		// Percentage scaling of font size.
			lf.lfHeight = ::MulDiv(lf.lfHeight, nScale, 100);
#endif
		CFont Font;
		VERIFY(Font.CreateFontIndirect(&lf));
		CFont *pOldFont = memDC.SelectObject(&Font);

		// Get the size of the text and adjust it if necessary.
		// If text is out of client area limits (scrolled), 
		//  it is clipped to the client area size.
		CRect rect;
		pWnd->GetClientRect(rect);
		CRect rcText(rect);
		UINT uFormat = DT_NOPREFIX;					// usually not a button or label text
		if (NULL != _tcschr(lpszText, _T('\n')))	// multiple lines of text
			uFormat |= DT_EDITCONTROL;				// don't draw partial last line
//		else
//			uFormat |= DT_SINGLELINE;
		memDC.DrawText(lpszText, -1, &rcText, uFormat | DT_CALCRECT);
		if (rcText.right > rect.right)				// limit the image size to the size of the control window
			rcText.right = rect.right;
		if (rcText.bottom > rect.bottom)
			rcText.bottom = rect.bottom;

		// Create bitmap and select it into the memory DC.
		CBitmap DragBitmap;
		if (DragBitmap.CreateCompatibleBitmap(&dc, rcText.Width(), rcText.Height()))
		{
			CBitmap *pOldBmp = memDC.SelectObject(&DragBitmap);
			COLORREF clrBk = bHiLightColor ? ::GetSysColor(COLOR_HIGHLIGHT) : dc.GetBkColor();
			COLORREF clrText = bHiLightColor ? ::GetSysColor(COLOR_HIGHLIGHTTEXT) : dc.GetTextColor();
			// Black color is used for transparency.
			// So choose closest color when text color is black.
			for (int n = 1; 0 == clrText; n++)
				clrText = dc.GetNearestColor(RGB(n, n, n));
			memDC.SetBkColor(clrBk);
			memDC.SetTextColor(clrText);
			VERIFY(memDC.DrawText(lpszText, -1, &rcText, uFormat));
			VERIFY(&DragBitmap == memDC.SelectObject(pOldBmp));
			bRet = SetDragImage(DragBitmap, pPoint, bHiLightColor ? 0 : clrBk);
		}
		VERIFY(&Font == memDC.SelectObject(pOldFont));
	}
	return bRet;
}

// Capture region from CWnd and use it as drag image.
// lpRect: The region in client coordinates. If NULL, the window's client area is used.
// nScale: Optional scaling.
//         If nScale < 0, the absolute value of nScale specifies the max. width or height of the image.
//         If nScale > 0, it is a percentage scaling factor.
//         If nScale is 0 or 100, no scaling will be performed.
// pPoint: Cursor position relative to left top corner of image. Centered if pPoint is NULL.
// clr:    Transparent color for the drag image.
//         When passing CLR_INVALID, the background color of the window is used.
bool COleDataSourceEx::InitDragImage(CWnd* pWnd, LPCRECT lpRect, int nScale, const CPoint *pPoint /*= NULL*/, COLORREF clr /*= CLR_INVALID*/)
{
	ASSERT_VALID(pWnd);

	bool bRet = false;
	if (pWnd && m_pDragSourceHelper)
	{
		CRect rect;
		if (lpRect)
			rect.CopyRect(lpRect);
		else
			pWnd->GetClientRect(rect);
		pWnd->ClientToScreen(rect);					// BitmapFromWnd() requires screen coordinates
		CBitmap Bitmap;
		COLORREF clrBk;
		// Copy region to Bitmap with optional scaling.
		// The background color is retrieved and optional passed as transparent color.
		if (BitmapFromWnd(pWnd, Bitmap, NULL, &rect, &clrBk, nScale))
		{
			HBITMAP hBitmap = (HBITMAP)Bitmap.Detach();
			if (CLR_INVALID == clr)					// use background color of window as transparent color
				clr = clrBk;
			if (0 != clr)							// if transparent color is not black
				hBitmap = ReplaceBlack(hBitmap);	// replace black pixels by RGB(1, 1, 1)
			if (hBitmap)
				bRet = SetDragImage(hBitmap, pPoint, clr);
		}
	}
	return bRet;
}

// Use a copy of the provided bitmap as drag image.
// nScale: Optional scaling.
//         If nScale < 0, the absolute value of nScale specifies the max. width or height of the image.
//         If nScale > 0, it is a percentage scaling factor.
// pPoint: Cursor position relative to left top corner of image. Centered if pPoint is NULL.
// clrBk:  The transparent color for the drag image (usually the background color of the bitmap).
//         With CLR_INVALID (-1), ::GetSysColor(COLOR_WINDOW) is used.
bool COleDataSourceEx::InitDragImage(HBITMAP hBitmap, int nScale, const CPoint *pPoint /*= NULL*/, COLORREF clrBk /*= CLR_INVALID*/)
{
	ASSERT(hBitmap);

	bool bRet = false;
	if (m_pDragSourceHelper && hBitmap)
	{
		if (CLR_INVALID == clrBk)
			clrBk = ::GetSysColor(COLOR_WINDOW);
		BITMAP bm;
		VERIFY(::GetObject(hBitmap, sizeof(bm), &bm));
		CSize sz(bm.bmWidth, bm.bmHeight);			// size of bitmap
		nScale = ScaleSize(nScale, sz);				// optional scaling
		// Replace black pixels by closest color.
		// Not supported for low color modes.
		// Not when transparent color is black.
		// When replacing black, create the copy as DIBSECTION
		//  to avoid creating of another copy by ReplaceBlack().
		bool bReplaceBlack = (bm.bmBitsPixel >= 16 && 0 != clrBk);
		HBITMAP hDragBitmap = (HBITMAP)::CopyImage(
			hBitmap, IMAGE_BITMAP, 
			sz.cx, sz.cy,
			bReplaceBlack ? LR_CREATEDIBSECTION : 0);
		if (hDragBitmap && bReplaceBlack)
			hDragBitmap = ReplaceBlack(hDragBitmap);
		if (hDragBitmap)
			bRet = SetDragImage(hDragBitmap, pPoint, clrBk);
	}
	return bRet;
}

// Get lowest bit of a DIBSECTION bitfields color mask.
// Helper function for ReplaceBlack().
DWORD COleDataSourceEx::GetLowestColorBit(DWORD dwBitField) const
{
	ASSERT(dwBitField);

	if (0 == dwBitField)
		return 0;
	DWORD dwMask = 1;
	while (0 == (dwBitField & 1))
	{
		dwBitField >>= 1;
		dwMask <<= 1;
	}
	return dwMask;
}

// Replace black by RGB(1, 1, 1) with 16, 24 and 32 BPP bitmaps.
// If passed bitmap is a DDB, a DIBSECTION copy is created and the original bitmap is destroyed.
// Don't call this when the transparent color is black.
// With drag images, transparent color pixels are replaced by black and
//  these pixels are used to generate a mask.
HBITMAP COleDataSourceEx::ReplaceBlack(HBITMAP hBitmap) const
{
	ASSERT(hBitmap);

	DIBSECTION ds;
	int nSize = ::GetObject(hBitmap, sizeof(ds), &ds);
	ASSERT(nSize);
	if (nSize && nSize < sizeof(ds) && ds.dsBm.bmBitsPixel >= 16)			
	{												// not a DIBSECTION
		// Create a DIBSECTION copy and delete the original bitmap.
		HBITMAP hCopy = (HBITMAP)::CopyImage(
			hBitmap, IMAGE_BITMAP, 0, 0,
			LR_CREATEDIBSECTION | LR_COPYDELETEORG);
		if (NULL != hCopy)
		{
			hBitmap = hCopy;
			nSize = ::GetObject(hBitmap, sizeof(ds), &ds);
		}
	}
	if (sizeof(ds) == nSize && ds.dsBm.bmBits && ds.dsBm.bmBitsPixel >= 16)			
	{												// is a high color DIBSECTION
		DWORD dwColor = 0x010101;					// default RGB(1, 1, 1) for 24/32 BPP
		// Get the RGB color masks with the lowest bit set and 
		//  combine them to the color that is closest to black.
		// Note that the dsBitfields members may be zero with 24 and 32 BPP bitmaps.
		// Then use the default 0x010101 value.
		if (ds.dsBitfields[0] && ds.dsBitfields[1] && ds.dsBitfields[2])
		{
			dwColor = GetLowestColorBit(ds.dsBitfields[0]) |
				GetLowestColorBit(ds.dsBitfields[1]) |
				GetLowestColorBit(ds.dsBitfields[2]);
		}
#ifdef _DEBUG
		else
			ASSERT(ds.dsBm.bmBitsPixel >= 24);		// must have dsBitfields with 16 bit colors
#endif
		LPBYTE pBits = static_cast<LPBYTE>(ds.dsBm.bmBits);
		if (16 == ds.dsBm.bmBitsPixel)
		{
			WORD wColor = static_cast<WORD>(dwColor);
			for (long i = 0; i < ds.dsBm.bmHeight; i++)
			{
				LPWORD pPixel = reinterpret_cast<LPWORD>(pBits);
				for (long j = 0; j < ds.dsBm.bmWidth; j++)
				{
					if (0 == *pPixel)
						*pPixel = wColor;
					++pPixel;
				}
				pBits += ds.dsBm.bmWidthBytes;		// begin of next/previous line
			}
		}
		else
		{
			int nBytePerPixel = ds.dsBm.bmBitsPixel / 8;
			BYTE ColorLow = static_cast<BYTE>(dwColor);
			BYTE ColorMid = static_cast<BYTE>(dwColor >> 8);
			BYTE ColorHigh = static_cast<BYTE>(dwColor >> 16);
			for (long i = 0; i < ds.dsBm.bmHeight; i++)
			{
				LPBYTE pPixel = pBits;
				for (long j = 0; j < ds.dsBm.bmWidth; j++)
				{
					if (0 == pPixel[0] && 0 == pPixel[1] && 0 == pPixel[2])
					{
						pPixel[0] = ColorLow;
						pPixel[1] = ColorMid;
						pPixel[2] = ColorHigh;
					}
					pPixel += nBytePerPixel;
				}
				pBits += ds.dsBm.bmWidthBytes;		// begin of next/previous line
			}
		}
	}
	return hBitmap;
}

//
// Caching functions
//

// Cache locale information.
// This is used with CF_TEXT and CF_OEMTEXT data for conversion when the target requests CF_UNICODETEXT.
// The LCID should be set to those of the text string.
// This is usually the LCID of the active thread which is used when passing zero.
// Passing LOCALE_USER_DEFAULT and LOCALE_SYSTEM_DEFAULT is supported.
bool COleDataSourceEx::CacheLocale(LCID nLCID /*= 0*/)
{
	HGLOBAL hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, sizeof(LCID));
	if (hGlobal)
	{
		LCID *lpLCID = static_cast<LCID *>(::GlobalLock(hGlobal));
		// Use thread locale when passing zero (LANG_NEUTRAL and SUBLANG_NEUTRAL).
		// Otherwise call ConvertDefaultLocale() to get the LCID when passing defaults.
        *lpLCID = (0 == LANGIDFROMLCID(nLCID)) ? ::GetThreadLocale() : ::ConvertDefaultLocale(nLCID);
		::GlobalUnlock(hGlobal);
		CacheGlobalData(CF_LOCALE, hGlobal);
	}
	return NULL != hGlobal;
}

// Cache text string for registered clipboard format.
bool COleDataSourceEx::CacheString(LPCTSTR lpszFormat, LPCTSTR lpszText, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(lpszFormat && *lpszFormat);

	CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(lpszFormat);
	return cfFormat ? CacheString(cfFormat, lpszText, dwTymed) : false;
}

// Cache text string.
bool COleDataSourceEx::CacheString(CLIPFORMAT cfFormat, LPCTSTR lpszText, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(cfFormat);
#ifdef _UNICODE
	ASSERT(cfFormat != CF_TEXT && cfFormat != CF_OEMTEXT && cfFormat != CF_DSPTEXT);
#else
	ASSERT(cfFormat != CF_UNICODETEXT);
#endif
	ASSERT(lpszText);
	ASSERT(dwTymed & (TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_FILE));

	size_t nSize = (_tcslen(lpszText) + 1) * sizeof(TCHAR);
	// Don't cache empty strings
	return (nSize > sizeof(TCHAR)) ? CacheFromBuffer(cfFormat, lpszText, nSize, dwTymed) : false;
}

// Convert Unicode string to multi byte and cache it as registered clipboard format.
bool COleDataSourceEx::CacheMultiByte(LPCTSTR lpszFormat, LPCWSTR lpszWide, UINT nCP /*= CP_THREAD_ACP*/, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(lpszFormat && *lpszFormat);

	CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(lpszFormat);
	return cfFormat ? CacheMultiByte(cfFormat, lpszWide, nCP, dwTymed) : false;
}

// Convert Unicode string to multi byte and cache it.
bool COleDataSourceEx::CacheMultiByte(CLIPFORMAT cfFormat, LPCWSTR lpszWide, UINT nCP /*= CP_THREAD_ACP*/, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(cfFormat);
	ASSERT(cfFormat != CF_UNICODETEXT);
	ASSERT(lpszWide);
	ASSERT(nCP != 1200);							// input is already UTF-16LE (use CacheString instead)
	ASSERT(nCP != 1201);							// WideCharToMultiByte() fails with UTF-16BE
	ASSERT(dwTymed & (TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_FILE));

	bool bRet = false;
	LPCSTR lpszMB = CDragDropHelper::WideCharToMultiByte(lpszWide, -1, nCP);
	if (lpszWide)
	{
		bRet = CacheFromBuffer(cfFormat, lpszMB, strlen(lpszMB) + 1, dwTymed);
		delete [] lpszMB;
	}
	return bRet;
}

#ifndef _UNICODE
// Convert multi byte string to Unicode and cache it.
bool COleDataSourceEx::CacheUnicode(CLIPFORMAT cfFormat, LPCSTR lpszAnsi, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(lpszAnsi && *lpszAnsi);

	bool bRet = false;
	LPCWSTR lpszWide = CDragDropHelper::MultiByteToWideChar(lpszText, -1);
	if (lpszWide)
	{
		bRet = CacheFromBuffer(cfFormat, lpszWide, (wcslen(lpszWide) + 1) * sizeof(WCHAR), dwTymed);
		delete [] lpszWide;
	}
	return bRet;
}
#endif

// Cache plain text.
// When providing clipboard data, the system performs implicit conversions between 
//  the CF_TEXT, CF_OEMTEXT and CF_UNICODETEXT formats when requesting data.
// But this does not apply to drag & drop operations.
// So we should pass data in all formats with drag & drop.
bool COleDataSourceEx::CacheText(LPCTSTR lpszText, bool bClipboard, bool bOem /*= false*/, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
#ifdef _UNICODE
	// Provide CF_UNICODETEXT
	bool bRet = CacheString(CF_UNICODETEXT, lpszText, dwTymed);
	if (bRet && !bClipboard)
	{
		// Provide CF_TEXT
		bool bCacheLocale = CacheMultiByte(CF_TEXT, lpszText, CP_THREAD_ACP, dwTymed);
		// Optional provide CF_OEMTEXT
		UINT nOemCP = (bCacheLocale && bOem) ? CDragDropHelper::GetCodePage(0, true) : 0;
		if (nOemCP)
			CacheMultiByte(CF_OEMTEXT, lpszText, nOemCP, dwTymed);
		// Provide locale information for ANSI text to Unicode conversion.
		if (bCacheLocale)
			CacheLocale();
	}
#else
	bool bRet = CacheString(CF_TEXT, lpszText, dwTymed);
	if (bRet && !bClipboard)
	{
		CacheUnicode(CF_UNICODETEXT, lpszText, dwTymed);
		UINT nOemCP = bOem ? CDragDropHelper::GetCodePage(0, true) : 0;
		if (nOemCP)
		{
			LPCSTR lpszOem = CDragDropHelper::MultiByteToMultiByte(lpszText, -1, CP_THREAD_ACP, nOemCP);
			if (lpszOem)
			{
				CacheString(CF_OEMTEXT, lpszOem, dwTymed);
				delete [] lpszOem;
			}
		}
	}
	if (bRet)
		CacheLocale();								// Provide locale information for ANSI text to Unicode conversion.
#endif
	return bRet;
}

// Uses a very simple text to RTF conversion.
bool COleDataSourceEx::CacheRtfFromText(LPCTSTR lpszText, CWnd *pWnd /*= NULL*/, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	CString strRtf;
	CTextToRTF Conv;
	bool bRet = (NULL != Conv.TextToRtf(strRtf, lpszText, pWnd));
	if (bRet)
		bRet = CacheRtf(strRtf.GetString(), dwTymed);
	return bRet;
}

// Cache HTML clipboard data.
// NOTES:
// - Input string must be valid HTML (using entities for reserved HTML chars '&"<>').
// - The input may optionally contain a HTML header and body tags.
//   If not present, they are added here (if pWnd is not NULL it is used to get font information).
//   If present they must met these requirements:
//   - If a html section is provided, head and body sections must be also present.
//   - The header should contain a charset definition for UTF-8:
//     <meta http-equiv="content-type" content="text/html; charset=utf-8">
// See "HTML Clipboard Format", http://msdn.microsoft.com/en-us/library/Aa767917.aspx
bool COleDataSourceEx::CacheHtml(LPCTSTR lpszHtml, CWnd *pWnd /*= NULL*/, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(lpszHtml && *lpszHtml);

	HGLOBAL hGlobal = NULL;
	CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(CF_HTML_STR);
	if (cfFormat)
	{
		CHtmlFormat HtmlFormat;
		hGlobal = HtmlFormat.SetHtml(lpszHtml, pWnd);
		if (hGlobal)
			CacheFromGlobal(cfFormat, hGlobal, HtmlFormat.GetSize(), dwTymed);
	}
	return NULL != hGlobal;
}

// Cache data as virtual files for drag & drop.
bool COleDataSourceEx::CacheVirtualFiles(unsigned nFiles, const VirtualFileData *pData, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(nFiles > 0);
	ASSERT(pData);
#ifdef _DEBUG
	for (unsigned i = 0; i < nFiles; i++)
		ASSERT(pData[i].IsValid());
#endif

	bool bRet = false;
	CLIPFORMAT cfContent = CDragDropHelper::RegisterFormat(CFSTR_FILECONTENTS);
	if (cfContent)
	{
		FORMATETC FormatEtc = { cfContent, NULL, DVASPECT_CONTENT, 0, dwTymed };
		for (unsigned i = 0; i < nFiles; i++)
		{
			// copy data to global memory
			HGLOBAL hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, pData[i].GetTotalSize());
			if (hGlobal)
			{
				LPBYTE lpBuf = static_cast<LPBYTE>(::GlobalLock(hGlobal));
				if (pData[i].GetBOMSize())
				{
					::CopyMemory(lpBuf, pData[i].GetBOM(), pData[i].GetBOMSize());
					lpBuf += pData[i].GetBOMSize();
				}
				::CopyMemory(lpBuf, pData[i].GetData(), pData[i].GetDataSize());
				::GlobalUnlock(hGlobal);
				if (CacheFromGlobal(cfContent, hGlobal, pData[i].GetTotalSize(), dwTymed, &FormatEtc))
					++FormatEtc.lindex;
			}
		}
		bRet = (FormatEtc.lindex == (int)nFiles);	// all files cached?
		if (bRet)
			bRet = CacheFileDescriptor(nFiles, pData);
	}
	return bRet;
}

// Cache file descriptors for virtual files.
bool COleDataSourceEx::CacheFileDescriptor(unsigned nFiles, const VirtualFileData *pData)
{
	ASSERT(nFiles > 0);
	ASSERT(pData);
#ifdef _DEBUG
	for (unsigned i = 0; i < nFiles; i++)
		ASSERT(pData[i].IsValid());
#endif

	bool bRet = false;
	CLIPFORMAT cfDescriptorA = CDragDropHelper::RegisterFormat(CFSTR_FILEDESCRIPTORA);
	CLIPFORMAT cfDescriptorW = CDragDropHelper::RegisterFormat(CFSTR_FILEDESCRIPTORW);
	if (cfDescriptorA && cfDescriptorW)
	{
		HGLOBAL hGlobalA = ::GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, 
			sizeof(FILEGROUPDESCRIPTORA) + (nFiles - 1) * sizeof(FILEDESCRIPTORA));
		HGLOBAL hGlobalW = ::GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, 
			sizeof(FILEGROUPDESCRIPTORW) + (nFiles - 1) * sizeof(FILEDESCRIPTORW));
		if (hGlobalA && hGlobalW)
		{
			bRet = true;
			LPFILEGROUPDESCRIPTORA pfgdA = static_cast<LPFILEGROUPDESCRIPTORA>(::GlobalLock(hGlobalA));
			LPFILEGROUPDESCRIPTORW pfgdW = static_cast<LPFILEGROUPDESCRIPTORW>(::GlobalLock(hGlobalW));
			pfgdA->cItems = pfgdW->cItems = nFiles;
			for (unsigned i = 0; i < nFiles; i++)
			{
				LPFILEDESCRIPTORA pfdA = &pfgdA->fgd[i];
				LPFILEDESCRIPTORW pfdW = &pfgdW->fgd[i];
				pfdA->dwFlags = pfdW->dwFlags =
					pData[i].GetTotalSize() ? FD_FILESIZE | FD_ATTRIBUTES : FD_ATTRIBUTES;
				pfdA->dwFileAttributes = pfdW->dwFileAttributes = FILE_ATTRIBUTE_NORMAL;
				pfdA->nFileSizeLow = pfdW->nFileSizeLow = static_cast<DWORD>(pData[i].GetTotalSize());
				LPCTSTR lpszFileName = pData[i].GetFileName();
#ifdef _UNICODE
				CStringA strA(lpszFileName);
				if (strncpy_s(pfdA->cFileName, sizeof(pfdA->cFileName), strA.GetString(), _TRUNCATE))
					bRet = false;
				if (wcsncpy_s(pfdW->cFileName, sizeof(pfdW->cFileName) / sizeof(pfdW->cFileName[0]), lpszFileName, _TRUNCATE))
					bRet = false;
#else
				CStringW strW(lpszFileName);
				if (strncpy_s(pfdA->cFileName, sizeof(pfdA->cFileName), lpszFileName, _TRUNCATE))
					bRet = false;
				if (wcsncpy_s(pfdW->cFileName, sizeof(pfdW->cFileName) / sizeof(pfdW->cFileName[0]), strW.GetString(), _TRUNCATE))
					bRet = false;
#endif
			}
			::GlobalUnlock(hGlobalW);
			::GlobalUnlock(hGlobalA);
		}
		if (bRet)
		{
			CacheGlobalData(cfDescriptorW, hGlobalW);
			CacheGlobalData(cfDescriptorA, hGlobalA);
		}
		else
		{
			if (NULL != hGlobalW)
				::GlobalFree(hGlobalW);
			if (NULL != hGlobalA)
				::GlobalFree(hGlobalA);
		}
	}
	return bRet;
}

// Cache bitmap.
// bClipboard: If true, only CF_DIBV5 is stored (other formats are provided by clipboard upon request). 
// hPal:       Optional palette to be cached as CF_PALETTE together with CF_BITMAP.
// dwTymed:    Storage medium for CF_DIBV5 and CF_DIB (CF_BITMAP is always TYMED_GDI).
//
// When providing clipboard data, the system performs implicit conversions between 
//  the CF_DIBV5, CF_DIB, CF_BITMAP, and CF_PALETTE formats when requesting data.
// But this does not apply to drag & drop operations.
// So we should pass data in all formats with drag & drop.
bool COleDataSourceEx::CacheBitmap(HBITMAP hBitmap, bool bClipboard, HPALETTE hPal /*= NULL*/, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	bool bRet = CacheDib(hBitmap, true, dwTymed);	// cache CF_DIBV5
	if (bRet && !bClipboard)
	{
		if (CacheDib(hBitmap, false, dwTymed))		// cache CF_DIB
			CacheDdb(hBitmap, hPal, true);			// cache CF_BITMAP and CF_PALETTE using copies of the objects
	}
	return bRet;
}

// Cache DDB (Device Dependant Bitmap) and an optional palette.
// Ownership of the GDI objects is passed to OLE. 
// Set bCopyObjects to true to pass copies of the objects.
bool COleDataSourceEx::CacheDdb(HBITMAP hBitmap, HPALETTE hPal, bool bCopyObjects)
{
	ASSERT(hBitmap);

	STGMEDIUM StgMedium;
	StgMedium.tymed = TYMED_GDI;					// type of medium is a GDI object
	StgMedium.pUnkForRelease = NULL;				// use default procedure to release storage
	StgMedium.hBitmap = bCopyObjects ? (HBITMAP)::OleDuplicateData(hBitmap, CF_BITMAP, 0) : hBitmap;
	if (StgMedium.hBitmap)
	{
		CacheData(CF_BITMAP, &StgMedium);			// cache bitmap
#if 0
		if (NULL == hPal)							// Use default palette if none passed.
		{
			BITMAP bm;
			if (sizeof(bm) == ::GetObject(hBitmap, sizeof(bm), &bm))
			{
				if (bm.bmBitsPixel <= 8)
				{
					hPal = static_cast<HPALETTE>(::GetStockObject(DEFAULT_PALETTE));
					bCopyObjects = false;
				}
			}
		}
#endif
		if (hPal)									// cache palette if provided
		{
			StgMedium.hBitmap = bCopyObjects ? (HBITMAP)::OleDuplicateData(hPal, CF_PALETTE, 0) : (HBITMAP)hPal;
			if (StgMedium.hBitmap)
				CacheData(CF_PALETTE, &StgMedium);
		}
	}
	return NULL != StgMedium.hBitmap;
}

// Create a device independant DIB from HBITMAP DDB or DIB and cache it.
//  hBitmap: Image. Ownership is unchanged (caller must release it).
//  bV5:     Create version 5 DIB if true.
bool COleDataSourceEx::CacheDib(HBITMAP hBitmap, bool bV5, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	size_t nSize = 0;
	HGLOBAL hDib = CreateDib(hBitmap, nSize, bV5, false);
	return hDib ? CacheFromGlobal(bV5 ? CF_DIBV5 : CF_DIB, hDib, nSize, dwTymed) : false;
}

// Cache bitmap as image.
// See RenderImage() for supported image types.
bool COleDataSourceEx::CacheImage(HBITMAP hBitmap, LPCTSTR lpszFormat, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(lpszFormat);
	return cfFormat ? CacheImage(hBitmap, cfFormat, dwTymed) : false;
}
  
// Cache bitmap as image.
// cfFormat defines the image type.
// See RenderImage() for supported image types.
bool COleDataSourceEx::CacheImage(HBITMAP hBitmap, CLIPFORMAT cfFormat, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(hBitmap);
	ASSERT(cfFormat);

	FORMATETC FormatEtc = { cfFormat, NULL, DVASPECT_CONTENT, -1, dwTymed };
	STGMEDIUM StgMedium;
	StgMedium.tymed = TYMED_NULL;
	bool bRet = (0 != RenderImage(hBitmap, &FormatEtc, &StgMedium));
	if (bRet)
		CacheData(cfFormat, &StgMedium, &FormatEtc);
	return bRet;
}
  
// Create a bitmap from text and cache it as CF_BITMAP.
bool COleDataSourceEx::CacheBitmapFromText(LPCTSTR lpszText, CWnd *pWnd)
{
	ASSERT(lpszText);
	ASSERT_VALID(pWnd);

	CBitmap Bitmap;
	bool bRet = BitmapFromText(Bitmap, pWnd, lpszText);
	if (bRet)
		bRet = CacheDdb((HBITMAP)Bitmap.Detach(), NULL, false);
	return bRet;
}

// Cache bitmap as virtual image file.
// hBitmap:      Bitmap.
// lpszFileName: Filename with extension but without path. If null, "Image.bmp" is used.
//               The extension defines the image format (BMP with unknown extensions).
bool COleDataSourceEx::CacheBitmapAsFile(HBITMAP hBitmap, LPCTSTR lpszFileName, DWORD dwTymed /*= TYMED_HGLOBAL*/)
{
	ASSERT(hBitmap);

	bool bRet = false;
	CLIPFORMAT cfContents = CDragDropHelper::RegisterFormat(CFSTR_FILECONTENTS);
	if (cfContents)
	{
		if (NULL == lpszFileName || _T('\0') == *lpszFileName)
			lpszFileName = _T("Image.bmp");
		GUID FileType = Gdiplus::ImageFormatBMP;
		LPCTSTR lpszExt = _tcsrchr(lpszFileName, _T('.'));
		if (lpszExt)
		{
			++lpszExt;
			if (0 == _tcsicmp(lpszExt, _T("tif")) || 0 == _tcsicmp(lpszExt, _T("tiff")))
				FileType = Gdiplus::ImageFormatTIFF;
			else if (0 == _tcsicmp(lpszExt, _T("jpg")) || 0 == _tcsicmp(lpszExt, _T("jpeg")))
				FileType = Gdiplus::ImageFormatJPEG;
			else if (0 == _tcsicmp(lpszExt, _T("gif")))
				FileType = Gdiplus::ImageFormatGIF;
			else if (0 == _tcsicmp(lpszExt, _T("png")))
				FileType = Gdiplus::ImageFormatPNG;
		}
		FORMATETC FormatEtc = { cfContents, NULL, DVASPECT_CONTENT, 0, dwTymed };
		STGMEDIUM StgMedium;
		StgMedium.tymed = TYMED_NULL;
		size_t nSize = RenderImage(hBitmap, &FormatEtc, &StgMedium, FileType);
		if (nSize)
		{
			VirtualFileData VirtualData;
			VirtualData.SetSize(nSize);
			VirtualData.SetFileName(lpszFileName);
			bRet = CacheFileDescriptor(1, &VirtualData);
			if (bRet)
				CacheData(cfContents, &StgMedium, &FormatEtc);
			else
				::ReleaseStgMedium(&StgMedium);
		}
	}
	return bRet;
}

#ifdef OBJ_SUPPORT
// Allocate and cache global memory block with serialized data.
bool COleDataSourceEx::CacheObjectData(CObject& Obj, LPCTSTR lpszFormat)
{
	ASSERT_VALID(&Obj);
	ASSERT(lpszFormat && *lpszFormat);

	HGLOBAL hGlobal = NULL;
	CLIPFORMAT cfFormat = CDragDropHelper::RegisterFormat(lpszFormat);
	if (cfFormat)
	{
		CSharedFile	sf(GMEM_MOVEABLE);
		CArchive ar(&sf, CArchive::store);
		Obj.Serialize(ar);
		ar.Close();
		hGlobal = sf.Detach();
		if (hGlobal)
			CacheGlobalData(cfFormat, hGlobal);	
	}
	return NULL != hGlobal;
}
#endif // OBJ_SUPPORT

// Cache CF_HDROP with a single file name.
bool COleDataSourceEx::CacheSingleFileAsHdrop(LPCTSTR lpszFileName)
{
	ASSERT(lpszFileName && *lpszFileName);

	size_t nNameSize = (_tcslen(lpszFileName) + 1) * sizeof(TCHAR);
	// DROPFILES followed by NULL terminated file name and an additional NULL char.
	HGLOBAL hGlobal = ::GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, 
		sizeof(DROPFILES) + sizeof(TCHAR) + nNameSize);
	if (hGlobal)
	{
		LPVOID lpData = ::GlobalLock(hGlobal);
		LPDROPFILES pDropFiles = static_cast<LPDROPFILES>(lpData);
		pDropFiles->pFiles = sizeof(DROPFILES);
#ifdef _UNICODE
		pDropFiles->fWide = 1;
#endif
		LPBYTE lpFiles = static_cast<LPBYTE>(lpData) + sizeof(DROPFILES);
		::CopyMemory(lpFiles, lpszFileName, nNameSize);
		::GlobalUnlock(hGlobal);
		CacheGlobalData(CF_HDROP, hGlobal);
	}
	return NULL != hGlobal;
}

// Cache data from memory using multiple storage types.
bool COleDataSourceEx::CacheFromBuffer(CLIPFORMAT cfFormat, LPCVOID lpData, size_t nSize, DWORD dwTymed, LPFORMATETC lpFormatEtc /*= NULL*/)
{
	ASSERT(cfFormat);
	ASSERT(lpData);
	ASSERT(nSize);
	ASSERT(dwTymed & (TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_FILE));

	FORMATETC FormatEtc = { cfFormat, NULL, DVASPECT_CONTENT, -1, dwTymed };
	if (NULL == lpFormatEtc)
		lpFormatEtc = &FormatEtc;
	else
	{
		lpFormatEtc->cfFormat = cfFormat;
		lpFormatEtc->tymed = dwTymed;
	}
	STGMEDIUM StgMedium;
	StgMedium.tymed = TYMED_NULL;
	bool bRet = RenderFromBuffer(lpData, nSize, NULL, 0, lpFormatEtc, &StgMedium);
	if (bRet)
		CacheData(cfFormat, &StgMedium, lpFormatEtc);
	return bRet;
}

// Cache global memory data using multiple storage types.
// The global memory is passed to OLE or released here.
bool COleDataSourceEx::CacheFromGlobal(CLIPFORMAT cfFormat, HGLOBAL hGlobal, size_t nSize, DWORD dwTymed, LPFORMATETC lpFormatEtc /*= NULL*/)
{
	ASSERT(cfFormat);
	ASSERT(hGlobal);
	ASSERT(GMEM_INVALID_HANDLE != ::GlobalFlags(hGlobal));
	ASSERT(nSize <= ::GlobalSize(hGlobal));
	ASSERT(dwTymed & (TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_FILE));

	FORMATETC FormatEtc = { cfFormat, NULL, DVASPECT_CONTENT, -1, dwTymed };
	if (NULL == lpFormatEtc)
		lpFormatEtc = &FormatEtc;
	else
	{
		lpFormatEtc->cfFormat = cfFormat;
		lpFormatEtc->tymed = dwTymed;
	}
	STGMEDIUM StgMedium;
	StgMedium.tymed = TYMED_NULL;
	bool bRet = RenderFromGlobal(hGlobal, nSize, lpFormatEtc, &StgMedium);
	if (bRet)
		CacheData(cfFormat, &StgMedium, lpFormatEtc);
	return bRet;
}

//
// Render support functions
//

BOOL COleDataSourceEx::OnRenderData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
	BOOL bRet = COleDataSource::OnRenderData(lpFormatEtc, lpStgMedium);
	// Use callback function to let the source window render the data.
	if (!bRet && m_pWnd && m_pOnRenderData)
		bRet = m_pOnRenderData(m_pWnd, this, lpFormatEtc, lpStgMedium);
	return bRet;
}

// Prepare STGMEDIUM with data from memory.
// lpData:      Pointer to data.
// nSize:       Size of data in bytes. Must be non zero.
// lpPrefix:    Optional prefix data (e.g. BOM or header structure).
// nPrefixSize: Size of optional prefix data. Must be zero if lpPrefix is NULL and non zero if lpPrefix is not NULL.
// lpFormatEtc: tymed member specifies requested storage types; other members are not used
// lpStgMedium: Filled here. Use passed value when rendering or set tymed to NULL when
//              preparing data for caching.
// NOTE: When lpFormatEtc->tymed is TYMED_FILE, lpStgMedium->lpszFileName must specify the file name.
bool COleDataSourceEx::RenderFromBuffer(LPCVOID lpData, size_t nSize, LPCVOID lpPrefix, size_t nPrefixSize, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
{
	ASSERT(lpData);
	ASSERT(nSize);
	ASSERT(NULL == lpPrefix || 0 != nPrefixSize);
	ASSERT(0 == nPrefixSize || NULL != lpPrefix);
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);
	ASSERT(lpFormatEtc->tymed & (TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_FILE));
	ASSERT(NULL != lpStgMedium->lpszFileName || 0 == (lpFormatEtc->tymed & TYMED_FILE));

	bool bRet = false;
	bool bNewMedium = true;
	HGLOBAL hGlobal = NULL;
	LPSTREAM lpStream = NULL;
	HANDLE hFile = INVALID_HANDLE_VALUE;
	if (NULL == lpPrefix)
		nPrefixSize = 0;

	// Use existing global memory: Fails if existing memory is not large enough.
	if (TYMED_HGLOBAL == lpStgMedium->tymed && NULL != lpStgMedium->hGlobal)
	{
		bNewMedium = false;
		if (nSize + nPrefixSize <= ::GlobalSize(lpStgMedium->hGlobal))
			hGlobal = lpStgMedium->hGlobal;
	}
	// Use existing stream: Rewind stream.
	else if (TYMED_ISTREAM == lpStgMedium->tymed && NULL != lpStgMedium->pstm)
	{
		bNewMedium = false;
		LARGE_INTEGER llNull = { 0, 0 };
		if (SUCCEEDED(lpStgMedium->pstm->Seek(llNull, STREAM_SEEK_SET, NULL)))
			lpStream = lpStgMedium->pstm;
	}
	// Use existing file: Open file (fails if file does not exist).
	else if (TYMED_FILE == lpStgMedium->tymed && NULL != lpStgMedium->lpszFileName)
	{
		bNewMedium = false;
		hFile = ::CreateFile(OLE2T(lpStgMedium->lpszFileName), GENERIC_WRITE, 0, NULL, TRUNCATE_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	}
	// Use new global memory: Allocate memory.
	else if (lpFormatEtc->tymed & TYMED_HGLOBAL)
	{
		hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, nSize + nPrefixSize);
	}
	// Use new stream: Create stream.
	else if (lpFormatEtc->tymed & TYMED_ISTREAM)
	{
		::CreateStreamOnHGlobal(NULL, TRUE, &lpStream);
	}
	// Use new file: Create file (fails if file exists).
	// NOTE: File name must be passed in lpStgMedium->lpszFileName as allocated OLE string!
	else if (lpFormatEtc->tymed & TYMED_FILE && NULL != lpStgMedium->lpszFileName)
	{
		// Using FILE_ATTRIBUTE_TEMPORARY may improve performance.
		// The system will not write the file to disk if enough cache memory is available.
		// This is especially useful with drag & drop operations where the file exists
		//  for short times only (it will be deleted when releasing the mouse button).
		hFile = ::CreateFile(OLE2T(lpStgMedium->lpszFileName), GENERIC_WRITE, 0, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL | FILE_ATTRIBUTE_TEMPORARY, NULL);
	}
	// Copy data to global memory
	if (hGlobal)
	{
		LPBYTE lpGlobal = static_cast<LPBYTE>(::GlobalLock(hGlobal));
		if (lpPrefix && nPrefixSize)
		{
			::CopyMemory(lpGlobal, lpPrefix, nPrefixSize);
			lpGlobal += nPrefixSize;
		}
		::CopyMemory(lpGlobal, lpData, nSize);
		::GlobalUnlock(hGlobal);
		if (bNewMedium)
		{
			lpStgMedium->tymed = TYMED_HGLOBAL;
			lpStgMedium->hGlobal = hGlobal;
			lpStgMedium->pUnkForRelease = NULL;
		}
		bRet = true;
	}
	// Write data to stream
	else if (lpStream)
	{
		bRet = true;
		ULARGE_INTEGER llSize;
		llSize.QuadPart = nSize + nPrefixSize;
		bRet = SUCCEEDED(lpStream->SetSize(llSize));
		if (bRet && nPrefixSize)
			bRet = SUCCEEDED(lpStream->Write(lpPrefix, static_cast<ULONG>(nPrefixSize), NULL));
		if (bRet)
			bRet = SUCCEEDED(lpStream->Write(lpData, static_cast<ULONG>(nSize), NULL));
		if (bNewMedium)
		{
			if (bRet)
			{
				lpStgMedium->tymed = TYMED_ISTREAM;
				lpStgMedium->pstm = lpStream;
				lpStgMedium->pUnkForRelease = NULL;
			}
			else
				lpStream->Release();
		}
	}
	// Write data to file
	else if (INVALID_HANDLE_VALUE != hFile)
	{
		bRet = true;
		DWORD dwWritten = 0;
		if (nPrefixSize)
			bRet = (0 != ::WriteFile(hFile, lpPrefix, static_cast<DWORD>(nPrefixSize), &dwWritten, NULL));
		if (bRet)
			bRet = (0 != ::WriteFile(hFile, lpData, static_cast<DWORD>(nSize), &dwWritten, NULL));
		::CloseHandle(hFile);
		if (bNewMedium)
		{
			if (bRet)
			{
				lpStgMedium->tymed = TYMED_FILE;
				lpStgMedium->pUnkForRelease = NULL;
			}
			else
			{
				::DeleteFile(OLE2T(lpStgMedium->lpszFileName));
				::CoTaskMemFree(lpStgMedium->lpszFileName);
				lpStgMedium->lpszFileName = NULL;
			}
		}
	}
	return bRet;
}

// Prepare STGMEDIUM with data provided as HGLOBAL.
// hGlobal:     Global memory. Passed to OLE or freed here.
// nSize:       Size of data in bytes. If zero, the global memory size is used.
// lpFormatEtc: tymed member specifies requested storage types; other members are not used.
// lpStgMedium: Filled here. Use passed value when rendering or set tymed to NULL when
//              preparing data for caching.
// NOTE: When lpFormatEtc->tymed is TYMED_FILE, lpStgMedium->lpszFileName must specify the file name.
bool COleDataSourceEx::RenderFromGlobal(HGLOBAL hGlobal, size_t nSize, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
{
	ASSERT(hGlobal);
	ASSERT(GMEM_INVALID_HANDLE != ::GlobalFlags(hGlobal));
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);
	ASSERT(lpFormatEtc->tymed & (TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_FILE));

	if (0 == nSize)
		nSize = ::GlobalSize(hGlobal);
	ASSERT(nSize);
	if (0 == nSize)
		return false;

	bool bRet = false;
	bool bFreeHGlobal = true;
	// Use existing global memory: Copy memory.
	if (TYMED_HGLOBAL == lpStgMedium->tymed && NULL != lpStgMedium->hGlobal)
	{
		bRet = (NULL != CDragDropHelper::CopyGlobalMemory(lpStgMedium->hGlobal, hGlobal, nSize));
	}
	// Use existing stream: Write content from global memory to stream.
	// Use existing file: Write content from global memory to file.
	else if ((TYMED_ISTREAM == lpStgMedium->tymed && NULL != lpStgMedium->pstm) ||
		(TYMED_FILE == lpStgMedium->tymed && NULL != lpStgMedium->lpszFileName))
	{
		bRet = RenderFromBuffer(::GlobalLock(hGlobal), nSize, NULL, 0, lpFormatEtc, lpStgMedium);
		::GlobalUnlock(hGlobal);
	}
	// Use new global memory: Pass handle to STGMEDIUM.
	else if (lpFormatEtc->tymed & TYMED_HGLOBAL)
	{
		lpStgMedium->tymed = TYMED_HGLOBAL;
		lpStgMedium->hGlobal = hGlobal;
		lpStgMedium->pUnkForRelease = NULL;
		bFreeHGlobal = false;
		bRet = true;
	}
	// Use new stream: Create stream using the global memory.
	else if (lpFormatEtc->tymed & TYMED_ISTREAM)
	{
		bRet = SUCCEEDED(::CreateStreamOnHGlobal(hGlobal, TRUE, &lpStgMedium->pstm));
		if (bRet)
		{
			bFreeHGlobal = false;					// memory is freed when the stream is released
			ULARGE_INTEGER llSize;
			llSize.QuadPart = nSize;
			if (FAILED(lpStgMedium->pstm->SetSize(llSize)))
			{
				lpStgMedium->pstm->Release();
				lpStgMedium->pstm = NULL;
			}
			else
			{
				lpStgMedium->tymed = TYMED_ISTREAM;
				lpStgMedium->pUnkForRelease = NULL;
			}
		}
	}
	// Use new file: Write content from global memory to file.
	else if (lpFormatEtc->tymed & TYMED_FILE)
	{
		lpStgMedium->lpszFileName = CreateOleString(CreateTempFileName().GetString());
		bRet = RenderFromBuffer(::GlobalLock(hGlobal), nSize, NULL, 0, lpFormatEtc, lpStgMedium);
		::GlobalUnlock(hGlobal);
	}
	// Free the global memory if it has not been passed to the STGMEDIUM or assigned to the stream.
	if (bFreeHGlobal)
		::GlobalFree(hGlobal);
	return bRet;
}

// Render plain text.
// Uses the passed cfFormat to define how the data is rendered while lpFormatEtc->cfFormat
//  defines the clipboard format. This allows rendering virtual file data according the 
//  format specified by cfFormat when lpFormatEtc->cfFormat is CFSTR_FILECONTENTS.
bool COleDataSourceEx::RenderText(CLIPFORMAT cfFormat, LPCTSTR lpszText, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
{
	bool bRet = false;
	switch (cfFormat)
	{ 
	case CF_UNICODETEXT :
		bRet = RenderUnicode(lpszText, lpFormatEtc, lpStgMedium);
		break;
	case CF_OEMTEXT :
		{
			UINT nOemCP = CDragDropHelper::GetCodePage(0, true);
			if (nOemCP)
				bRet = RenderMultiByte(lpszText, lpFormatEtc, lpStgMedium, nOemCP);
		}
		break;
	case CF_BITMAP :
	case CF_DIB :
	case CF_DIBV5 :
		// The control window pointer is required to get the DC and the font.
		if (m_pWnd)
		{
			CBitmap Bitmap;
			if (BitmapFromText(Bitmap, m_pWnd, lpszText))
			{
				if (CF_BITMAP == lpFormatEtc->cfFormat)
				{
					bRet = (TYMED_NULL == lpStgMedium->tymed);
					if (bRet)
					{
						lpStgMedium->tymed = TYMED_GDI;
						lpStgMedium->hBitmap = (HBITMAP)Bitmap.Detach();
						lpStgMedium->pUnkForRelease = NULL;
					}
				}
				else
					bRet = RenderBitmap((HBITMAP)Bitmap, lpFormatEtc, lpStgMedium);
			}
		}
		break;
	default :
		if (CDragDropHelper::IsRegisteredFormat(cfFormat, CF_HTML_STR))
			bRet = RenderHtml(lpszText, lpFormatEtc, lpStgMedium, m_pWnd);
		// All other formats are treated as ANSI text (including ASCII text like with RTF formats).
		else //if (CF_TEXT == cfFormat || CF_DSPTEXT == cfFormat || cfFormat >= 0xC000)
			bRet = RenderMultiByte(lpszText, lpFormatEtc, lpStgMedium);
	}
	return bRet;
}

// Render multi byte string.
// The passed string is TCHAR* and will be converted to multi byte with Unicode builds.
bool COleDataSourceEx::RenderMultiByte(LPCTSTR lpszText, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, UINT nCP /*= CP_THREAD_ACP*/) const
{
	ASSERT(lpszText && *lpszText);
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);
	ASSERT(CF_UNICODETEXT != lpFormatEtc->cfFormat);

	bool bRet = false;
#ifdef _UNICODE
	LPCSTR lpszMultiByte = CDragDropHelper::WideCharToMultiByte(lpszText, -1, nCP);
#else
	LPCSTR lpszMultiByte = NULL;
	// No conversion necessary with ASCII text or 
	//  when nCP is identical to the effective code page of the current thread.
	if (20127 == nCP || CP_THREAD_ACP == nCP || CDragDropHelper::CompareCodePages(CP_THREAD_ACP, nCP))
		lpszMultiByte = lpszText;
	else
		lpszMultiByte = CDragDropHelper::MultiByteToMultiByte(lpszText, -1, nCP);
#endif
	if (lpszMultiByte)
	{
		size_t nBomLen = 0;
		LPCSTR lpszBOM = NULL;
		// Write BOM when creating a virtual UTF-8 file. 
		if (CP_UTF8 == nCP && 
			CDragDropHelper::IsRegisteredFormat(lpFormatEtc->cfFormat, CFSTR_FILECONTENTS))
		{
			nBomLen = 3;
			lpszBOM = "\xEF\xBB\xBF";
		}
		bRet = RenderFromBuffer(
			lpszMultiByte, 
			strlen(lpszMultiByte) + 1, 
			lpszBOM, nBomLen,
			lpFormatEtc, lpStgMedium);
#ifndef _UNICODE
		if (lpszMultiByte != lpszText)
#endif
			delete [] lpszMultiByte;
	}
	return bRet;
}

// Render string as Unicode text.
bool COleDataSourceEx::RenderUnicode(LPCTSTR lpszText, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
{
	ASSERT(lpszText && *lpszText);
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);

	size_t nBomLen = 0;
	LPCSTR lpszBOM = NULL;
	// Write BOM when storing as virtual file.
	if (CDragDropHelper::IsRegisteredFormat(lpFormatEtc->cfFormat, CFSTR_FILECONTENTS))
	{
		nBomLen = 2;
		lpszBOM = "\xFE\xFF";
	}
#ifdef _UNICODE
	return RenderFromBuffer(
		lpszText, (wcslen(lpszText) + 1) * sizeof(WCHAR), 
		lpszBOM, nBomLen, 
		lpFormatEtc, lpStgMedium);
#else
	bool bRet = false;
	LPCWSTR lpszWide = CDragDropHelper::MultiByteToWideChar(lpszText, -1);
	if (lpszWide)
	{
		bRet = RenderFromBuffer(
			lpszWide, (wcslen(lpszWide) + 1) * sizeof(WCHAR), 
			lpszBOM, nBomLen, 
			lpFormatEtc, lpStgMedium);
		delete [] lpszWide;
	}
	return bRet;
#endif
}

// Render HTML clipboard format or virtual HTML file from HTML content string.
// The passed string must be valid HTML.
// If pWnd is not NULL, it is used to determine the font and add font styles with the HTML clipboard format.
bool COleDataSourceEx::RenderHtml(LPCTSTR lpszHtml, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, CWnd *pWnd) const
{
	ASSERT(lpszHtml);
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);

	bool bRet = false;
	// HTML clipboard format (ASCII header followed by the UTF-8 encoded HTML content).
	if (CDragDropHelper::IsRegisteredFormat(lpFormatEtc->cfFormat, CF_HTML_STR))
	{
		CHtmlFormat Conv;
		HGLOBAL hGlobal = Conv.SetHtml(lpszHtml, pWnd);
		if (hGlobal)
			bRet = RenderFromGlobal(hGlobal, Conv.GetSize(), lpFormatEtc, lpStgMedium);
	}
	// Virtual HTML file (UTF-8 encoded with BOM). 
	else if (CDragDropHelper::IsRegisteredFormat(lpFormatEtc->cfFormat, CFSTR_FILECONTENTS))
	{
		bRet = RenderMultiByte(lpszHtml, lpFormatEtc, lpStgMedium, CP_UTF8);
	}
	return bRet;
}

// Render bitmap as CF_BITMAP, CF_DIB, CF_DIBV5, BMP image file content, or MIME type.
bool COleDataSourceEx::RenderBitmap(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
{
	bool bRet = false;
	switch (lpFormatEtc->cfFormat)
	{
	case CF_BITMAP :
	case CF_PALETTE :
		bRet = RenderDdb(hBitmap, lpFormatEtc, lpStgMedium);
		break;
	case CF_DIB :
	case CF_DIBV5 :
		bRet = RenderDib(hBitmap, lpFormatEtc, lpStgMedium);
		break;
	default :
		// Convert bitmap according to lpFormatEtc->cfFormat.
		// With CFSTR_FILECONTENTS, a bitmap file image is stored.
		bRet = (0 != RenderImage(hBitmap, lpFormatEtc, lpStgMedium));
	}
	return bRet;
}

// Render bitmap as CF_BITMAP or palette as CF_PALETTE.
// hBitmap: HBITMAP or HPALETTE
// lpFormatEtc->cfFormat must be CF_BITMAP or CF_PALETTE
// lpStgMedium->tymed must be NULL (can't perform in place filling with GDI objects)
bool COleDataSourceEx::RenderDdb(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
{
	ASSERT(hBitmap);
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);
	ASSERT(CF_BITMAP == lpFormatEtc->cfFormat || CF_PALETTE == lpFormatEtc->cfFormat);
	ASSERT(TYMED_NULL == lpStgMedium->tymed);

	HBITMAP hCopy = NULL;
	// Can't perform in place filling with GDI objects.
	if (TYMED_NULL == lpStgMedium->tymed)
	{
		// GDI object will be owned by OLE. So we must create a copy.
		hCopy = (HBITMAP)::OleDuplicateData(hBitmap, lpFormatEtc->cfFormat, 0);
		if (NULL != hCopy)
		{
			lpStgMedium->tymed = TYMED_GDI;
			lpStgMedium->hBitmap = hCopy;
			lpStgMedium->pUnkForRelease = NULL;
		}
	}
	return NULL != hCopy;
}

// Render bitmap as CF_DIB or CF_DIBV5
bool COleDataSourceEx::RenderDib(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
{
	ASSERT(hBitmap);
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);
	ASSERT(CF_DIB == lpFormatEtc->cfFormat || CF_DIBV5 == lpFormatEtc->cfFormat);

	size_t nSize = 0;
	// Convert DDB to DIB stored in gobal memory
	HGLOBAL hGlobal = CreateDib(hBitmap, nSize, CF_DIBV5 == lpFormatEtc->cfFormat, false);
	return hGlobal ? RenderFromGlobal(hGlobal, nSize, lpFormatEtc, lpStgMedium) : false;
}

// Render bitmap as image.
// The image type is defined by lpFormatEtc->cfFormat.
// This uses CImage::Save() accepting an IStream to store converted images.
// Returns the size in bytes upon success and 0 on errors.
size_t COleDataSourceEx::RenderImage(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
{
	ASSERT(hBitmap);
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);

	GUID FileType = GUID_NULL;
	if (CF_TIFF == lpFormatEtc->cfFormat)
		FileType = Gdiplus::ImageFormatTIFF;
	// This requires special handling by the other RenderImage() function.
//	else if (CF_ENHMETAFILE == lpFormatEtc->cfFormat)
//		FileType = Gdiplus::ImageFormatEMF;
//	else if (CF_DIB == lpFormatEtc->cfFormat) // || CF_DIBV5 == lpFormatEtc->cfFormat || CF_BITMAP == lpFormatEtc->cfFormat)
//		FileType = Gdiplus::ImageFormatMemoryBMP;
	else
	{
		TCHAR lpszName[128];
		if (::GetClipboardFormatName(lpFormatEtc->cfFormat, lpszName, sizeof(lpszName) / sizeof(lpszName[0])) > 0)
		{
			if (0 == _tcscmp(lpszName, _T("image/bmp")) || 
				0 == _tcscmp(lpszName, _T("image/x-windows-bmp")) ||
				0 == _tcsicmp(lpszName, _T("Windows Bitmap")) ||
				0 == _tcsicmp(lpszName, CFSTR_FILECONTENTS))
			{
				FileType = Gdiplus::ImageFormatBMP;
			}
			else if (0 == _tcscmp(lpszName, _T("image/png")) || 
				0 == _tcsicmp(lpszName, _T("PNG")))
			{
				FileType = Gdiplus::ImageFormatPNG;
			}
			else if (0 == _tcscmp(lpszName, _T("image/gif")) || 
				0 == _tcsicmp(lpszName, _T("GIF")))
			{
				FileType = Gdiplus::ImageFormatGIF;
			}
			else if (0 == _tcscmp(lpszName, _T("image/jpeg")) || 
				0 == _tcsicmp(lpszName, _T("JPEG")))
			{
				FileType = Gdiplus::ImageFormatJPEG;
			}
			else if (0 == _tcscmp(lpszName, _T("image/tiff")) || 
				0 == _tcsicmp(lpszName, _T("TIFF")))
			{
				FileType = Gdiplus::ImageFormatTIFF;
			}
		}
	}
	return GUID_NULL == FileType ? 0 : RenderImage(hBitmap, lpFormatEtc, lpStgMedium, FileType);
}

// Render bitmap as image.
// The image type is defined by the passed GDI+ file type GUID.
// Note that the data object will be a bitmap file beginning with a BITMAPFILEHEADER 
//  with the GUID Gdiplus::ImageFormatBMP.
// This uses CImage::Save() accepting an IStream to store converted images.
// Returns the size in bytes upon success and 0 on errors.
size_t COleDataSourceEx::RenderImage(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, const GUID& FileType) const
{
	ASSERT(hBitmap);
	ASSERT(lpFormatEtc);
	ASSERT(lpStgMedium);
	ASSERT(GUID_NULL != FileType);
	ASSERT(lpFormatEtc->tymed & (TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_FILE));

	bool bHGlobal = false;
	bool bNewMedium = true;
	size_t nSize = 0;
	IStream *lpStream = NULL;
	CString strFileName;

	// Use existing global memory: Create temporary stream; memory is freed when releasing the stream.
	if (TYMED_HGLOBAL == lpStgMedium->tymed && NULL != lpStgMedium->hGlobal)
	{
		::CreateStreamOnHGlobal(NULL, TRUE, &lpStream);
		bNewMedium = false;
		bHGlobal = true;
	}
	// Use existing stream: Rewind.
	else if (TYMED_ISTREAM == lpStgMedium->tymed && NULL != lpStgMedium->pstm)
	{
		LARGE_INTEGER llNull = { 0, 0 };
		if (SUCCEEDED(lpStgMedium->pstm->Seek(llNull, STREAM_SEEK_SET, NULL)))
			lpStream = lpStgMedium->pstm;
		bNewMedium = false;
	}
	// Use existing file: Get file name.
	else if (TYMED_FILE == lpStgMedium->tymed && NULL != lpStgMedium->lpszFileName)
	{
		strFileName = lpStgMedium->lpszFileName;
		bNewMedium = false;
	}
	// Use new stream: Create stream.
	else if (lpFormatEtc->tymed & TYMED_ISTREAM)
	{
		::CreateStreamOnHGlobal(NULL, TRUE, &lpStream);
	}
	// Use new global memory: Create temporary stream; memory is not freed when releasing the stream.
	else if (lpFormatEtc->tymed & TYMED_HGLOBAL)
	{
		::CreateStreamOnHGlobal(NULL, FALSE, &lpStream);
		bHGlobal = true;
	}
	// Use new file: Create temp file name with image specific extension.
	else if (lpFormatEtc->tymed & TYMED_FILE)
	{
		LPCTSTR lpszExt = _T("bmp");
		if (Gdiplus::ImageFormatPNG == FileType)
			lpszExt = _T("png");
		else if (Gdiplus::ImageFormatGIF == FileType)
			lpszExt = _T("gif");
		if (Gdiplus::ImageFormatTIFF == FileType)
			lpszExt = _T("tif");
		if (Gdiplus::ImageFormatJPEG == FileType)
			lpszExt = _T("jpg");
		strFileName = CreateTempFileName(lpszExt);
	}
	// Write image to file and get the size.
	if (!strFileName.IsEmpty())
	{
		CImage Image;
		Image.Attach(hBitmap);
		if (SUCCEEDED(Image.Save(strFileName.GetString(), FileType)))
		{
			WIN32_FILE_ATTRIBUTE_DATA fi;
			if (::GetFileAttributesEx(strFileName.GetString(), GetFileExInfoStandard, &fi))
			{
				nSize = fi.nFileSizeLow;
				if (bNewMedium)
				{
					lpStgMedium->lpszFileName = CreateOleString(strFileName.GetString());
					if (lpStgMedium->lpszFileName)
					{
						lpStgMedium->tymed = TYMED_FILE;
						lpStgMedium->pUnkForRelease = NULL;
					}
					else
						nSize = 0;
				}
			}
			if (0 == nSize && bNewMedium)
				::DeleteFile(strFileName.GetString());
		}
		// Must Detach. Otherwise the bitmap is deleted by the CImage destructor.
		Image.Detach();
	}
	// Write image to stream and get the size.
	else if (lpStream)
	{
		CImage Image;
		Image.Attach(hBitmap);
		if (SUCCEEDED(Image.Save(lpStream, FileType)))
		{
			STATSTG st;
			if (SUCCEEDED(lpStream->Stat(&st, STATFLAG_NONAME)))
				nSize = static_cast<size_t>(st.cbSize.QuadPart);
		}
		// Must Detach. Otherwise the bitmap is deleted by the CImage destructor.
		Image.Detach();
		// Get the global memory handle from the stream with global memory objects.
		HGLOBAL hGlobal = NULL;
		if (bHGlobal)
			::GetHGlobalFromStream(lpStream, &hGlobal);
		if (0 == nSize)
		{
			// Release stream upon errors if not an existing one.
			if (bNewMedium || bHGlobal)
				lpStream->Release();
			lpStream = NULL;
			// Free the global memory handle with new global memory objects.
			// This has not been freed by releasing the stream.
			if (bNewMedium && NULL != hGlobal)
				::GlobalFree(hGlobal);
		}

		// Global memory: Use the HGLOBAL of the temporary stream.
		if (lpStream && bHGlobal)
		{
			if (NULL == hGlobal)
				nSize = 0;
			else
			{
				if (bNewMedium)
				{
					lpStgMedium->tymed = TYMED_HGLOBAL;
					lpStgMedium->hGlobal = hGlobal;
					lpStgMedium->pUnkForRelease = NULL;
				}
				// This fails when the existing memory is too small for the image.
				else if (NULL == CDragDropHelper::CopyGlobalMemory(lpStgMedium->hGlobal, hGlobal, nSize))
					nSize = 0;
			}
			// Release the temporary stream.
			// The memory is also freed when not passed to STGMEDIUM.
			lpStream->Release();
		}
		// New stream: Fill STGMEDIUM.
		else if (lpStream && bNewMedium)
		{
			lpStgMedium->tymed = TYMED_ISTREAM;
			lpStgMedium->pstm = lpStream;
			lpStgMedium->pUnkForRelease = NULL;
		}
	}
	return nSize;
}

//
// Bitmap support functions
//

// Create DIB in global memory.
// nSize: Set to the real total size upon success (allocated size may be larger).
// bV5:   If true create a V5 DIB (BITMAPV5HEADER, color table, image bits).
// bFile: If true create a BMP file image (BITMAPFILEHEADER, BITMAPINFOHEADER, color table, image bits).
HGLOBAL COleDataSourceEx::CreateDib(HBITMAP hBitmap, size_t& nSize, bool bV5, bool bFile) const
{
	ASSERT(hBitmap);

	HGLOBAL hGlobal = NULL;
	HBITMAP hDib = NULL;
	DIBSECTION ds;

	// Create a DIBSECTION copy of the bitmap if necessary.
	int nSizeDS = ::GetObject(hBitmap, sizeof(ds), &ds);
	if (sizeof(ds) == nSizeDS)						// hBitmap is already a DIB section bitmap
		hDib = hBitmap;
	else if (nSizeDS >= sizeof(BITMAP))				// hBitmap is a DDB
	{
		hDib = (HBITMAP)::CopyImage(hBitmap, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
		if (hDib)
			nSizeDS = ::GetObject(hDib, sizeof(ds), &ds);
	}
	if (hDib && sizeof(ds) == nSizeDS && ds.dsBm.bmBits)
	{
		// No version 5 DIB with bitmap file
		if (bFile)
			bV5 = false;
		// Get header size
		size_t nHdrSize = bV5 ? sizeof(BITMAPV5HEADER) : ds.dsBmih.biSize;
		if (bFile)
			nHdrSize += sizeof(BITMAPFILEHEADER);
		if (ds.dsBmih.biBitCount <= 8)
		{
			if (0 == ds.dsBmih.biClrUsed)
				ds.dsBmih.biClrUsed = 1 << (DWORD)ds.dsBmih.biBitCount;
			nHdrSize += ds.dsBmih.biClrUsed * sizeof(RGBQUAD);
		}
		else
		{
			ds.dsBmih.biClrUsed = 0;				// no optimize color table
			if (BI_BITFIELDS == ds.dsBmih.biCompression)
				nHdrSize += sizeof(ds.dsBitfields);
		}
		// This should be not necessary.
		// But many MSDN examples perform these calculations.
		if (0 == ds.dsBmih.biSizeImage)
		{
			if (0 == ds.dsBm.bmWidthBytes)
				ds.dsBm.bmWidthBytes = ((ds.dsBm.bmWidth * ds.dsBm.bmBitsPixel + 31) & ~31) / 8;
			ds.dsBmih.biSizeImage = (DWORD)(ds.dsBm.bmHeight * ds.dsBm.bmWidthBytes);
		}
		hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, nHdrSize + ds.dsBmih.biSizeImage);
		if (hGlobal)
		{
			// Pass total size to caller
			nSize = nHdrSize + ds.dsBmih.biSizeImage;
			LPBYTE lpData = static_cast<LPBYTE>(::GlobalLock(hGlobal));
			LPBYTE lpDst = lpData;

			// Initialize file header
			if (bFile)
			{
				LPBITMAPFILEHEADER pBmHdr = reinterpret_cast<LPBITMAPFILEHEADER>(lpDst);
				pBmHdr->bfType = 0x4D42; // "BM"
				pBmHdr->bfSize = static_cast<DWORD>(nSize);
				pBmHdr->bfReserved1 = pBmHdr->bfReserved2 = 0;
				pBmHdr->bfOffBits = static_cast<DWORD>(nHdrSize);
				lpDst += sizeof(BITMAPFILEHEADER);
			}
			// Copy info header
			::CopyMemory(lpDst, &ds.dsBmih, ds.dsBmih.biSize);
			// Extend info header to V5 header
			if (bV5)
			{
				// Clear the additional members
				::ZeroMemory(lpDst + sizeof(BITMAPINFOHEADER), sizeof(BITMAPV5HEADER) - sizeof(BITMAPINFOHEADER));
				// Set the different and additional members
				LPBITMAPV5HEADER pHdrV5 = reinterpret_cast<LPBITMAPV5HEADER>(lpDst);
				pHdrV5->bV5Size = sizeof(BITMAPV5HEADER);
				pHdrV5->bV5CSType = LCS_sRGB;
				if (BI_BITFIELDS == ds.dsBmih.biCompression)
				{
					pHdrV5->bV5RedMask = ds.dsBitfields[0];
					pHdrV5->bV5GreenMask = ds.dsBitfields[1];
					pHdrV5->bV5BlueMask = ds.dsBitfields[2];
				}
				lpDst += sizeof(BITMAPV5HEADER);
			}
			else
				lpDst += ds.dsBmih.biSize;
			// Copy optional color masks
			if (BI_BITFIELDS == ds.dsBmih.biCompression && ds.dsBmih.biBitCount > 8)
			{
				::CopyMemory(lpDst, ds.dsBitfields, sizeof(ds.dsBitfields));
				lpDst += sizeof(ds.dsBitfields);
			}
			// Copy color table
			if (ds.dsBmih.biClrUsed)
				VERIFY(ds.dsBmih.biClrUsed == GetDibColorTable(hDib, ds.dsBmih.biClrUsed, reinterpret_cast<RGBQUAD*>(lpDst)));
			// Copy the image bits
			::CopyMemory(lpData + nHdrSize, ds.dsBm.bmBits, ds.dsBmih.biSizeImage);
			::GlobalUnlock(hGlobal);
		}
	}
	if (hDib && hDib != hBitmap)
		::DeleteObject(hDib);						// delete the copy
	return hGlobal;
}

UINT COleDataSourceEx::GetDibColorTable(HBITMAP hBitmap, UINT nColors, RGBQUAD* pColors) const
{
	if (nColors)
	{
		HDC hScreen = ::CreateCompatibleDC(NULL);
		HBITMAP hOld = (HBITMAP)::SelectObject(hScreen, hBitmap);
		nColors = ::GetDIBColorTable(hScreen, 0, nColors, pColors);
		VERIFY(hBitmap == (HBITMAP)::SelectObject(hScreen, hOld));
		VERIFY(::DeleteDC(hScreen));
	}
	return nColors;
}

// Create bitmap from text.
bool COleDataSourceEx::BitmapFromText(CBitmap& Bitmap, CWnd* pWnd, LPCTSTR lpszText) const
{
	ASSERT(NULL == Bitmap.m_hObject);				// bitmap handle attached to CBitmap
	ASSERT_VALID(pWnd);
	ASSERT(lpszText);

	bool bRet = false;
	CDC memDC;
	CClientDC dc(pWnd);
	VERIFY(memDC.CreateCompatibleDC(&dc));

	// Use the font of the control window.
	CFont *pOldFont = memDC.SelectObject(pWnd->GetFont());

	// Calculate the image size
	CRect rc(0, 0, 0, 0);
	memDC.DrawText(lpszText, -1, &rc, DT_NOPREFIX | DT_CALCRECT);

	// Create bitmap and select it into a memory DC.
	if (Bitmap.CreateCompatibleBitmap(&dc, rc.Width(), rc.Height()))
	{
		CBitmap *pOldBmp = memDC.SelectObject(&Bitmap);
		memDC.SetBkColor(dc.GetBkColor());
		memDC.SetTextColor(dc.GetTextColor());
		// Erase the background
		memDC.FillSolidRect(&rc, memDC.GetBkColor());
		// Draw the text. Using DT_NOCLIP is faster.
		VERIFY(memDC.DrawText(lpszText, -1, &rc, DT_NOPREFIX | DT_NOCLIP));
		VERIFY(&Bitmap == memDC.SelectObject(pOldBmp));
		bRet = true;
	}
	memDC.SelectObject(pOldFont);
	return bRet;
}

// Get CWnd window content into bitmap.
// lpRect can be used to specify a specific area in screen coordinates.
//  If lpRect is NULL, the client area is copied (excluding caption and frames).
//  To copy the complete window including caption and frames, pass lpRect filled by
//   calling pWnd->GetWindowRect().
// nScale allows optional scaling of the image.
//  If > 0, it is a percentage scaling value (> 100 enlarges the image and < 100 reduces the size).
//  If < 0, the image is resized to the absolute number of pixels if width or height are bigger.
//  If nScale is 0 or 100, no resizing occurs.
// Provided as static function to be used for other purposes like writing window content
//  to a bitmap file.
/*static*/ BOOL COleDataSourceEx::BitmapFromWnd(CWnd *pWnd, CBitmap &Bitmap, CPalette *pPal, LPCRECT lpRect /*= NULL*/, COLORREF *lpBkClr /*= NULL*/, int nScale /*= 0*/)
{
	ASSERT_VALID(pWnd);

	// The WindowDC covers the complete CWnd including caption and frames.
	// When lpRect is NULL, we must adjust the rect to cover the client area only.
	// When lpRect is not NULL, we must convert screen coordinates to window relative coordinates.
	CWindowDC dc(pWnd);								// Use CWindowDC to allow access to non-client area
	if (lpBkClr)									// Optional return background color
		*lpBkClr = dc.GetBkColor();
	CRect rectSrc, rcAdj;
	pWnd->GetWindowRect(rcAdj);						// get pos and size of window in screen coordinates
	if (lpRect)
		rectSrc.CopyRect(lpRect);					// passed screen coordinates
	else
	{
		pWnd->GetClientRect(rectSrc);				// get size of client area
		pWnd->ScreenToClient(rcAdj);				// convert wnd rect to client coordinates to get left/top offset
	}
	rectSrc.OffsetRect(-rcAdj.left, -rcAdj.top);	// convert to window (not client) relative coordinates

	CDC memDC;										// memory DC to receive the bitmap
	VERIFY(memDC.CreateCompatibleDC(&dc));
	// Optional scale the captured bitmap if the device supports stretching.
	CSize sizeDst = rectSrc.Size();
	nScale = ScaleSize(nScale, sizeDst, memDC.GetSafeHdc());
	BOOL bRet = Bitmap.CreateCompatibleBitmap(&dc, sizeDst.cx, sizeDst.cy);
	if (bRet)
	{
		CBitmap* pOldBitmap = memDC.SelectObject(&Bitmap);
		if (nScale)
		{
			bRet = memDC.StretchBlt(
				0, 0, sizeDst.cx, sizeDst.cy, 
				&dc, 
				rectSrc.left, rectSrc.top, rectSrc.Width(), rectSrc.Height(), 
				SRCCOPY);
		}
		else
		{
			bRet = memDC.BitBlt(
				0, 0, sizeDst.cx, sizeDst.cy, 
				&dc, rectSrc.left, rectSrc.top, SRCCOPY);
		}
		VERIFY(&Bitmap == memDC.SelectObject(pOldBitmap));
		// Optional create a logical palette if the device supports a palette.
		if (bRet && pPal && (dc.GetDeviceCaps(RASTERCAPS) & RC_PALETTE))
		{
			// The LOGPALETTE struct contains one PALETTEENTRY element to represent the entries.
			// So add space for up to 255 additional entries.
			BYTE pBuf[sizeof(LOGPALETTE) + 255 * sizeof(PALETTEENTRY)]; 
			LPLOGPALETTE pLP = reinterpret_cast<LPLOGPALETTE>(pBuf);
			pLP->palVersion = 0x300;
			pLP->palNumEntries = static_cast<WORD>(::GetSystemPaletteEntries(dc.GetSafeHdc(), 0, 256, pLP->palPalEntry));
			if (pLP->palNumEntries)
				pPal->CreatePalette(pLP);
		}
	}
	return bRet;
}

// Get window content into bitmap using WM_PRINTCLIENT.
// This copies the complete client area even when the window is (partially) hidden.
// This will not work with custom windows that did not support WM_PRINTCLIENT.
/*static*/ BOOL COleDataSourceEx::BitmapFromWnd(CWnd *pWnd, CBitmap &Bitmap)
{
	ASSERT_VALID(pWnd);

	CDC memDC;
	CClientDC dc(pWnd);
	VERIFY(memDC.CreateCompatibleDC(&dc));
	CRect rc;
	pWnd->GetClientRect(rc);
	BOOL bRet = Bitmap.CreateCompatibleBitmap(&dc, rc.Width(), rc.Height());
	if (bRet)
	{
		CBitmap* pOldBitmap = memDC.SelectObject(&Bitmap);
		memDC.FillSolidRect(&rc, dc.GetBkColor());
		pWnd->SendMessage(WM_PRINTCLIENT, (WPARAM)memDC.GetSafeHdc(), 
			PRF_CLIENT | PRF_ERASEBKGND | PRF_CHILDREN | PRF_OWNED | PRF_NONCLIENT);
		VERIFY(&Bitmap == memDC.SelectObject(pOldBitmap));
	}
	return bRet;
}

// Get window content into bitmap using PrintWindow().
// This will not work with all windows.
/*static*/ BOOL COleDataSourceEx::PrintWndToBitmap(CWnd *pWnd, CBitmap &Bitmap, bool bClient)
{
	ASSERT_VALID(pWnd);

	CDC memDC;
	CWindowDC dc(pWnd);
	VERIFY(memDC.CreateCompatibleDC(&dc));
	CRect rc;
	if (bClient)
		pWnd->GetClientRect(rc);
	else
	{
		pWnd->GetWindowRect(rc);
		rc.OffsetRect(-rc.left, -rc.top);
	}
	BOOL bRet = Bitmap.CreateCompatibleBitmap(&dc, rc.Width(), rc.Height());
	if (bRet)
	{
		CBitmap* pOldBitmap = memDC.SelectObject(&Bitmap);
		memDC.FillSolidRect(&rc, dc.GetBkColor());
		pWnd->PrintWindow(&memDC, bClient ? PW_CLIENTONLY : 0);
		VERIFY(&Bitmap == memDC.SelectObject(pOldBitmap));
	}
	return bRet;
}

// Adjust CSize according to scaling value.
// Used by static function BitmapFromWnd().
// Returns input value nScale if size has been changed and zero if not.
/*static*/ int COleDataSourceEx::ScaleSize(int nScale, CSize& size, HDC hDC /*= NULL*/)
{
	if (100 == nScale ||							// no scaling with 100% and if screen does not support stretching
		(hDC && !::GetDeviceCaps(hDC, RASTERCAPS & RC_STRETCHBLT)))
		nScale = 0;
	else if (nScale > 0)							// percentage scaling value
	{
		size.cx = ::MulDiv(size.cx, nScale, 100);
		size.cy = ::MulDiv(size.cy, nScale, 100); 
	}
	else if (nScale < 0)							// nScale specifies max. width and height
	{
		int nMax = -nScale;
		if (size.cx > nMax && size.cx >= size.cy)	// width exceeds the limit and width >= height
		{
			size.cy = ::MulDiv(size.cy, nMax, size.cx);
			size.cx = nMax;
		}
		else if (size.cy > nMax)					// height exceeds the limit
		{
			size.cx = ::MulDiv(size.cx, nMax, size.cy);
			size.cy = nMax;
		}
		else										// width and height are smaller than the limit
			nScale = 0;								//  no scaling
	}
	return nScale;
}

//
// Helper functions
//

// Create temp file path and name.
// If lpszExt is not NULL, the default TMP extension is replaced.
CString COleDataSourceEx::CreateTempFileName(LPCTSTR lpszExt /*= NULL*/) const
{
	TCHAR lpszTempPath[MAX_PATH - 14];
	TCHAR lpszTempFile[MAX_PATH];
	CString str;

	// NOTE: This may return the temp folder of the current user which may be not accessible for other users!
	DWORD dwTempLen = ::GetTempPath(sizeof(lpszTempPath) / sizeof(lpszTempPath[0]), lpszTempPath);
	if (dwTempLen && dwTempLen < sizeof(lpszTempPath) / sizeof(lpszTempPath[0]))
	{
		if (::GetTempFileName(lpszTempPath, _T("ODS"), 0, lpszTempFile))
		{
			str = lpszTempFile;
			if (lpszExt && *lpszExt)
			{
				int nPos = str.ReverseFind(_T('.'));
				if (nPos)
				{
					str.Truncate(nPos + 1);
					if (_T('.') == *lpszExt)
						++lpszExt;
					str += lpszExt;
				}
			}
		}
	}
	return str;
}

// Copy string to allocated OLE memory. 
LPOLESTR COleDataSourceEx::CreateOleString(LPCTSTR lpszStr) const
{
	ASSERT(lpszStr);

	LPOLESTR lpszOle = NULL;
	if (lpszStr && *lpszStr)
	{
#ifdef _UNICODE
		SIZE_T nSize = wcslen(lpszStr) + 1;
		lpszOle = static_cast<LPOLESTR>(::CoTaskMemAlloc(nSize * sizeof(WCHAR)));
		if (lpszOle)
			wcscpy_s(lpszOle, nSize, lpszStr);
#else
		CStringW strW(lpszStr);
		SIZE_T nSize = strW.GetLength() + 1;
		lpszOle = static_cast<LPOLESTR>(::CoTaskMemAlloc(nSize * sizeof(WCHAR)));
		if (lpszOle)
			wcscpy(lpszOle, nSize, strW.GetString());
#endif
	}
	return lpszOle;
}

// Drag source helper support.
// See http://www.codeproject.com/Articles/3530/DragSourceHelper-MFC
// Interface map for our object.
BEGIN_INTERFACE_MAP(COleDataSourceEx, COleDataSource)
	INTERFACE_PART(COleDataSourceEx, IID_IDataObject, DataObjectEx)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) COleDataSourceEx::XDataObjectEx::AddRef()
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) COleDataSourceEx::XDataObjectEx::Release()
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->ExternalRelease();
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return (HRESULT)pThis->ExternalQueryInterface(&iid, ppvObj);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::GetData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.GetData(lpFormatEtc, lpStgMedium);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::GetDataHere(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.GetDataHere(lpFormatEtc, lpStgMedium);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::QueryGetData(LPFORMATETC lpFormatEtc)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.QueryGetData(lpFormatEtc);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::GetCanonicalFormatEtc(LPFORMATETC lpFormatEtcIn, LPFORMATETC lpFormatEtcOut)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.GetCanonicalFormatEtc(lpFormatEtcIn, lpFormatEtcOut);
}

// IDragSourceHelper calls this with custom clipboard formats.
// This call fails because custom formats are not supported by the COleDataSource::XDataObject::SetData() function.
// So cache the data here to make them available for displaying the drag image by the IDropTargetHelper.
// Checking for the format numbers here will not work because these numbers may differ.
// We may check for the format names here, but this requires getting and comparing the names which
//  may also change with Windows versions.
// Names are: DragImageBits, DragContext, IsShowingLayered, DragWindow, IsShowingText
STDMETHODIMP COleDataSourceEx::XDataObjectEx::SetData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, BOOL bRelease)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	HRESULT hr = pThis->m_xDataObject.SetData(lpFormatEtc, lpStgMedium, bRelease);
	if (DATA_E_FORMATETC == hr &&					// Error: Invalid FORMATETC structure
		(lpFormatEtc->tymed & (TYMED_HGLOBAL | TYMED_ISTREAM)) &&
		lpFormatEtc->cfFormat >= 0xC000)
	{
		pThis->CacheData(lpFormatEtc->cfFormat, lpStgMedium, lpFormatEtc);
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::EnumFormatEtc(DWORD dwDirection, LPENUMFORMATETC* ppenumFormatetc)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.EnumFormatEtc(dwDirection, ppenumFormatetc);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::DAdvise(LPFORMATETC lpFormatEtc, DWORD advf, LPADVISESINK pAdvSink, LPDWORD pdwConnection)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.DAdvise(lpFormatEtc, advf, pAdvSink, pdwConnection);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::DUnadvise(DWORD dwConnection)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.DUnadvise(dwConnection);
}

STDMETHODIMP COleDataSourceEx::XDataObjectEx::EnumDAdvise(LPENUMSTATDATA* ppenumAdvise)
{
	METHOD_PROLOGUE(COleDataSourceEx, DataObjectEx)
	return pThis->m_xDataObject.EnumDAdvise(ppenumAdvise);
}

#ifdef _DEBUG
void COleDataSourceEx::AssertValid() const
{
	COleDataSource::AssertValid();
	ASSERT(NULL == m_pOnRenderData || NULL != m_pWnd);
}

void COleDataSourceEx::Dump(CDumpContext& dc) const
{
	COleDataSource::Dump(dc);

	dc << "m_bSetDescriptionText = " << m_bSetDescriptionText;
	dc << "\nm_nDragResult = " << m_nDragResult;

	dc << "\n";
}
#endif //_DEBUG

