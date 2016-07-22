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

// OLE Drag & Drop target support class.
#pragma once

#include "DragDropHelper.h"

// User defined messages for drag events.
#define WM_APP_DRAG_ENTER	(WM_APP + 1)			// WPARAM: COleDataObject*, LPARAM: DragParameters* 
#define WM_APP_DRAG_LEAVE	(WM_APP + 2)			// WPARAM: unused,          LPARAM: unused 
#define WM_APP_DRAG_OVER	(WM_APP + 3)			// WPARAM: COleDataObject*, LPARAM: DragParameters* 
#define WM_APP_DROP_EX		(WM_APP + 4)			// WPARAM: COleDataObject*, LPARAM: DragParameters* 
#define WM_APP_DROP			(WM_APP + 5)			// WPARAM: COleDataObject*, LPARAM: DragParameters* 
#define WM_APP_DRAG_SCROLL	(WM_APP + 6)			// WPARAM: dwKeyState,      LPARAM: point 
#define WM_APP_DO_SCROLL	(WM_APP + 7)			// WPARAM: vertical steps,  LPARAM: horizonal steps 

// Auto scroll mode flags
#define AUTO_SCROLL_HORZ_BAR	0x00001				// Auto scroll when over horizontal scroll bar
#define AUTO_SCROLL_VERT_BAR	0x00002				// Auto scroll when over vertical scroll bar
#define AUTO_SCROLL_INSET_HORZ	0x00004				// Auto scroll when over horizontal inset region
#define AUTO_SCROLL_INSET_VERT	0x00008				// Auto scroll when over vertical inset region
#define AUTO_SCROLL_INSET_BAR	0x00010				// Auto scroll when over inset region and bar visible
#define AUTO_SCROLL_MASK		0x0FFFF
#define AUTO_SCROLL_DEFAULT		0x10000				// Use default handling (no OnDragScroll() handler required)

#define AUTO_SCROLL_BARS		(AUTO_SCROLL_HORZ_BAR | AUTO_SCROLL_VERT_BAR)
#define AUTO_SCROLL_INSETS		(AUTO_SCROLL_INSET_HORZ | AUTO_SCROLL_INSET_VERT)

// Callback function prototypes
typedef DROPEFFECT	(CALLBACK* OnDragEnterFunc)(CWnd*, COleDataObject*, DWORD, CPoint);
typedef DROPEFFECT	(CALLBACK* OnDragOverFunc)(CWnd*, COleDataObject*, DWORD, CPoint);
typedef DROPEFFECT	(CALLBACK* OnDropExFunc)(CWnd*, COleDataObject*, DROPEFFECT, DROPEFFECT, CPoint);
typedef BOOL		(CALLBACK* OnDropFunc)(CWnd*, COleDataObject*, DROPEFFECT, CPoint);
typedef void		(CALLBACK* OnDragLeaveFunc)(CWnd*);    
typedef DROPEFFECT	(CALLBACK* OnDragScrollFunc)(CWnd*, DWORD, CPoint);
typedef void	    (CALLBACK* ScrollFunc)(CWnd*, int, int);

class COleDropTargetEx : public COleDropTarget
{
// Construction
public:
	COleDropTargetEx();

	typedef struct tagDragParameters				// data passed with drag event messages
	{
		DWORD			dwKeyState;					// WM_APP_DRAG_ENTER, WM_APP_DRAG_OVER
		CPoint			point;
		DROPEFFECT		dropEffect;					// dropDefault parameter with WM_APP_DROP_EX
		DROPEFFECT		dropList;					// with WM_APP_DROP_EX only
	} DragParameters;

	// Overrides
	virtual DROPEFFECT OnDragEnter(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	virtual DROPEFFECT OnDragOver(CWnd* pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	virtual DROPEFFECT OnDropEx(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point);
	virtual BOOL OnDrop(CWnd* pWnd, COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point);
	virtual void OnDragLeave(CWnd* pWnd);    
	virtual DROPEFFECT OnDragScroll(CWnd* pWnd, DWORD dwKeyState, CPoint point);

	// Implementation
public:
	virtual ~COleDropTargetEx();
#ifdef _DEBUG
	virtual void Dump(CDumpContext& dc) const;
#endif

	inline void SetUseMsg()							{ m_bUseMsg = true; }
	inline bool GetUseDropDescription() const		{ return m_bUseDropDescription; }
	inline void SetScrollMode(unsigned nMode, ScrollFunc p = NULL)	
	{ m_nScrollMode = nMode; m_pScrollFunc = p; }

	// Pass callback functions.
	inline void SetOnDragEnter(OnDragEnterFunc p)	{ m_pOnDragEnter = p; }
	inline void SetOnDragOver(OnDragOverFunc p)		{ m_pOnDragOver = p; }
	inline void SetOnDropEx(OnDropExFunc p)			{ m_pOnDropEx = p; }
	inline void SetOnDrop(OnDropFunc p)				{ m_pOnDrop = p; }
	inline void SetOnDragLeave(OnDragLeaveFunc p)	{ m_pOnDragLeave = p; }
	inline void SetOnDragScroll(OnDragScrollFunc p)	{ m_pOnDragScroll = p; }

	DROPEFFECT GetDropEffect(DWORD dwKeyState, DROPEFFECT dwDefault = DROPEFFECT_MOVE) const;
	DROPEFFECT FilterDropEffect(DROPEFFECT dwEffect) const;
	// Preferred drop effect specified by the drag source.
	// If this it not DROPEFFECT_NONE, it should be used by the target
	//  in the drop effect selection decision.
	inline DROPEFFECT GetPreferredDropEffect() const { return m_nPreferredDropEffect; }
	// True when drop description text is shown by the drag source.
	inline bool IsTextAllowed() const { return m_bTextAllowed; }
	// True when drop descriptions can be used (Vista or later and visual styles enabled).
	inline bool CanShowDescription() const { return m_bCanShowDescription; }
	// Drop effects passed to DoDragDrop by drop source
	inline DROPEFFECT GetDropEffects() const { return m_nDropEffects; }

	bool SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert = NULL);
	bool SetDropInsertText(LPCWSTR lpszDropInsert);
	bool SetDropDescription(DROPIMAGETYPE nImageType, LPCWSTR lpszText, bool bCreate);
	bool SetDropDescription(DROPEFFECT dwEffect);
	inline bool ClearDropDescription()
	{ return SetDropDescription(DROPIMAGE_INVALID, NULL, false); }

	// COleDataObject access functions.

	// Standard text formats
	UINT GetCodePage(COleDataObject* pDataObject, CLIPFORMAT cfFormat = CF_TEXT) const;
	CLIPFORMAT GetBestTextFormat(COleDataObject* pDataObject) const;
	LPTSTR GetString(CLIPFORMAT cfFormat, COleDataObject* pDataObject, UINT nCP = 0) const;
	LPTSTR GetText(COleDataObject* pDataObject) const;
	LPSTR  GetHtmlUtf8String(COleDataObject* pDataObject, bool bAll) const;
	LPTSTR GetHtmlString(COleDataObject* pDataObject, bool bAll) const;
	LPCSTR GetHtmlData(COleDataObject* pDataObject, size_t& nSize) const;

	size_t GetCharCount(COleDataObject* pDataObject, unsigned *pnLines) const;
	size_t GetCharCount(CLIPFORMAT cfFormat, COleDataObject* pDataObject, unsigned *pnLines, UINT nCP = 0) const;

	// CF_HDROP
	CString GetSingleFileName(COleDataObject* pDataObject) const;
	unsigned GetFileNameCount(COleDataObject* pDataObject) const;
	LPTSTR GetFileListAsText(COleDataObject* pDataObject) const;
	size_t GetFileListAsTextLength(COleDataObject* pDataObject, unsigned *pnLines) const;

	// Bitmaps
	HBITMAP GetBitmap(COleDataObject* pDataObject) const;
	HPALETTE GetPalette(COleDataObject* pDataObject) const;
	HBITMAP GetBitmapFromDDB(COleDataObject* pDataObject) const;
	HBITMAP GetBitmapFromDIB(COleDataObject* pDataObject) const;
	HBITMAP GetBitmapFromImage(COleDataObject* pDataObject) const;
	HBITMAP GetBitmapFromImage(COleDataObject* pDataObject, CLIPFORMAT cfFormat) const;
	HBITMAP GetDragImageBitmap(COleDataObject* pDataObject) const;

	LPVOID GetData(COleDataObject* pDataObject, CLIPFORMAT cfFormat, size_t& nSize, LPFORMATETC lpFormatEtc = NULL) const;
	HGLOBAL GetDataAsGlobal(COleDataObject* pDataObject, CLIPFORMAT cfFormat, size_t& nSize, LPFORMATETC lpFormatEtc /*= NULL*/) const;

protected:
	void		ClearStateFlags();
	void		OnPostDrop(COleDataObject* pDataObject, DROPEFFECT dropEffect, CPoint point);
	DROPEFFECT	DoDefaultScroll(CWnd *pWnd, DWORD dwKeyState, CPoint point);
	void		SetScrollPos(const CWnd* pWnd, int nStep, int nBar) const;

	inline bool IsScrollBarVisible(const SCROLLBARINFO& sbi) const { return STATE_SYSTEM_INVISIBLE != sbi.rgstate[0]; }
	inline bool IsScrollBarInvisible(const SCROLLBARINFO& sbi) const { return STATE_SYSTEM_INVISIBLE == sbi.rgstate[0]; }
	inline bool IsScrollBarActive(const SCROLLBARINFO& sbi) const { return 0 == sbi.rgstate[0]; }
	inline bool CanScrollUpRight(const SCROLLBARINFO& sbi) const { return 0 == sbi.rgstate[2]; }
	inline bool CanScrollDownLeft(const SCROLLBARINFO& sbi) const { return 0 == sbi.rgstate[4]; }
	// Check if scrolling can be performed when over the inset region.
	// AUTO_SCROLL_INSET_BAR: True if scroll bar visible and enabled, and can scroll into direction.
	// Otherwise: True if scroll bar invisible or can scroll into direction.
	inline bool CanScrollInsetUpRight(const SCROLLBARINFO& sbi) const
	{ 
		return (m_nScrollMode & AUTO_SCROLL_INSET_BAR) ? (IsScrollBarActive(sbi) && CanScrollUpRight(sbi)) :
			(IsScrollBarInvisible(sbi) || CanScrollUpRight(sbi)); 
	}
	inline bool CanScrollInsetDownLeft(const SCROLLBARINFO& sbi) const
	{ 
		return (m_nScrollMode & AUTO_SCROLL_INSET_BAR) ? (IsScrollBarActive(sbi) && CanScrollDownLeft(sbi)) :
			(IsScrollBarInvisible(sbi) || CanScrollDownLeft(sbi)); 
	}

	size_t GetFileListAsTextLength(HDROP hDrop, unsigned& nLines) const;

protected:
	bool				m_bUseMsg;					// Call window handlers by sending messages
	unsigned			m_nScrollMode;				// Auto scroll handling
	IDropTargetHelper*	m_pDropTargetHelper;		// Drag image helper
	OnDragEnterFunc		m_pOnDragEnter;				// Callback function pointers
	OnDragOverFunc		m_pOnDragOver;
	OnDropExFunc		m_pOnDropEx;
	OnDropFunc			m_pOnDrop;
	OnDragLeaveFunc		m_pOnDragLeave;
	OnDragScrollFunc	m_pOnDragScroll;
	ScrollFunc			m_pScrollFunc;
	CDropDescription	m_DropDescription;			// User defined default drop description text
private:
	bool				m_bCanShowDescription;		// Set when drop descriptions can be used (Vista or later)
	bool				m_bUseDropDescription;		// If true, drop descriptions are updated by this class
	bool				m_bDescriptionUpdated;		// Internal flag to detect if drop description has been updated
	bool				m_bEntered;					// Set when entered the target. Use by OnDragScroll().
	bool				m_bHasDragImage;			// Cached "DragWindow" data object exists state
	bool				m_bTextAllowed;				// Cached "DragSourceHelperFlags" data object
	DROPEFFECT			m_nPreferredDropEffect;		// Cached "Preferred DropEffect" data object
	DROPEFFECT			m_nDropEffects;				// Drop effects passed to DoDragDrop

// Interface Maps
public:
	BEGIN_INTERFACE_PART(DropTargetEx, IDropTarget)
		INIT_INTERFACE_PART(COleDropTargetEx, DropTargetEx)
		STDMETHOD(DragEnter)(LPDATAOBJECT, DWORD, POINTL, LPDWORD);
		STDMETHOD(DragOver)(DWORD, POINTL, LPDWORD);
		STDMETHOD(DragLeave)();
		STDMETHOD(Drop)(LPDATAOBJECT, DWORD, POINTL, LPDWORD);
	END_INTERFACE_PART(DropTargetEx)

	DECLARE_INTERFACE_MAP()

};
