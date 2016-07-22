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

#pragma once
#include "resource.h"								// symbols
#include "OleDropTargetEx.h"

// Support for dropping files names from Explorer.
// If set, m_nDropFiles specifies the behaviour:
//  0   = don't accept file names from Explorer
//  -1  = accept any number of file names if total length is less than max length
//  > 0 = accept up to this number of file names if total length is less than max length
// If not defined, dropping of file names from the explorer is handled by the CEdit
//  base class when the style bit WS_EX_ACCEPTFILES is set.
#define MYEDIT_DROP_FILES							// HDROP support implemented

class COleDataSourceEx;

// CMyEdit

class CMyEdit : public CEdit
{
	DECLARE_DYNAMIC(CMyEdit)

public:
					CMyEdit();
	virtual			~CMyEdit();

	void			Init(LPCTSTR lpszText = NULL, UINT nLimit = 0);
	bool			HasSelection() const;
	bool			CanSelectAll() const;
	inline bool		CanEdit() const			{ return (IsWindowEnabled() && !(GetStyle() & ES_READONLY)); }
	inline bool		CanCopy() const			{ return HasSelection(); }
	inline bool		CanCopyAll() const		{ return (GetWindowTextLength() > 0); }
	inline bool		CanDeselect() const		{ return HasSelection(); }
	inline bool		CanPaste() const		{ return CanEdit() && ::IsClipboardFormatAvailable(CF_TEXT); }	// also true with CF-UNICODETEXT
	inline bool		CanCut() const			{ return CanEdit() && HasSelection(); }
	inline bool		CanClear() const		{ return CanEdit() && GetWindowTextLength() > 0; }

	inline bool		GetOleMove() const		{ return m_bOleMove; }
	inline bool		GetOleDropSelf() const	{ return m_bOleDropSelf; }

	inline void		SetDragImageType(int n, bool b)	{ m_nDragImageType = n; m_bAllowDropDesc = b; }
	inline void		SetDragImageScale(int n)	{ m_nDragImageScale = n; }
	inline void		SetAllowDragMove(bool b)	{ m_bOleMove = b; }
	inline void		SetAllowDropSelf(bool b)	{ m_bOleDropSelf = b; }
	inline void		SetSelectDropped(bool b)	{ m_bSelectDropped = b; }

	bool GetSelectionRect(CRect& rect, const CPoint *pPoint = NULL);
	void GetSelectedText(CString& str) const;
	void PasteAnsiText(CLIPFORMAT cfFormat, UINT nCP = CP_ACP);
	inline void PasteAnsiText(LPCTSTR lpszFormat, UINT nCP = CP_ACP)
	{ PasteAnsiText(static_cast<CLIPFORMAT>(::RegisterClipboardFormat(lpszFormat)), nCP); }

private:
	bool			m_bDragging;					// set when control is OLE source (detect if dropping on source)
	bool			m_bOleMove;						// allow moving of content from control
	bool			m_bOleDropSelf;					// allow Drag&Drop inside control
	bool			m_bSelectDropped;				// select drpped text upon insertion
	bool			m_bAllowDropDesc;				// allow drop description text (Vista and later)
	int				m_nDragImageType;
	int				m_nDragImageScale;
	int				m_nDropLen;						// cached value with total size of drop source in TCHARS
	COleDropTargetEx	m_DropTarget;				// OLE Drag&Drop target support

protected:
#ifdef MYEDIT_DROP_FILES							// HDROP handled by this class
	int				m_nDropFiles;					// accept file names from explorer (number of names, -1 = all)
#endif

	void PasteHtml(bool bAll);
	COleDataSourceEx* CacheOleData(const CString& str, bool bClipboard);

//	virtual void PreSubclassWindow();

	// OLE Drag&Drop target handlers.
	LRESULT	OnDragEnter(WPARAM wParam, LPARAM lParam);
	LRESULT	OnDragOver(WPARAM wParam, LPARAM lParam);
	LRESULT	OnDropEx(WPARAM wParam, LPARAM lParam);
	LRESULT	OnDragScroll(WPARAM wParam, LPARAM lParam);
	LRESULT	OnDoScroll(WPARAM wParam, LPARAM lParam);

	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);

public:
	afx_msg void OnEditCopy();
	afx_msg void OnEditCopyAll();
	afx_msg void OnEditSelectAll();
	afx_msg void OnEditDeselect();
	afx_msg void OnEditUndo();
	afx_msg void OnEditPaste();
	afx_msg void OnEditCut();
	afx_msg void OnEditClear();
	afx_msg void OnEditPasteCsv();
	afx_msg void OnEditPasteHtml();
	afx_msg void OnEditPasteHtmlAll();
	afx_msg void OnEditPasteRtf();
	afx_msg void OnEditPasteFileNames();
	afx_msg void OnEditPasteFileContent();

	DECLARE_MESSAGE_MAP()
};
