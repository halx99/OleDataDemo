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

#include "OleDropTargetEx.h"

class COleDataSourceEx;

// MyListCtrl.h:
//
//  Report list control providing data for clipboard and OLE Drag&Drop.

// GetBitmap() options.
#define GET_BITMAP_HILIGHT	0x01		// selected items are drawn using system highlight colors
#define GET_BITMAP_GRID		0x02		// draw grid lines
#define GET_BITMAP_BORDER	0x04		// draw border around list (1 pixel wide solid black line)
#define GET_BITMAP_HDRFRAME	0x08		// draw header items as buttons

// Helper class for text copy.
struct CGetTextHelper
{
	CGetTextHelper();
	~CGetTextHelper();
	void Clear();

	inline bool IsInitialized() const 
	{ return NULL != m_pWidth && NULL != m_pOrder; }
	inline int GetItemLines() const 
	{ return m_bAll ? m_nItemCount : m_nSelectedCount; }
	inline int GetTotalLines() const 
	{ return GetItemLines() + m_bHeader; }
	inline int GetTotalCells() const 
	{ return m_nUsedColumns * GetTotalLines(); }
	inline int GetTextSize() const 
	{ return m_nTextSize; }
	inline int GetTextSizeLineFeed() const 
	{ return m_nTextSize + 2 * GetTotalLines(); }
	inline int GetTotalColumnSeparatorCount() const 
	{ return m_nUsedColumns > 1 ? (m_nUsedColumns - 1) * GetTotalLines() : 0; }

	bool m_bAll;
	bool m_bHeader;
	bool m_bHidden;
	bool m_bHasCheckBoxes;
	int m_nItemCount;
	int m_nColumnCount;
	int m_nSelectedCount;
	int m_nUsedColumns;
	int m_nTextSize;
	int *m_pWidth;
	int *m_pOrder;
	int *m_pFormat;
};

/////////////////////////////////////////////////////////////////////////////
// CMyListCtrl 

class CMyListCtrl : public CListCtrl
{
	DECLARE_DYNAMIC(CMyListCtrl)
public:
	CMyListCtrl();
//	virtual ~CMyListCtrl();

	void	Init();
	bool	GetBitmap(CBitmap& Bitmap, bool bAll = false, bool bHeader = false, unsigned nFlags = 0);
	void	GetText(CString& strTxt, CString& strCsv, CString& strHtml, bool bAll = false, bool bHeader = false, bool bHidden = false);
	int		GetPlainText(CString& str, bool bAll, bool bHeader, bool bHidden);
	int		GetCsvText(CString& strCsv, bool bAll, bool bHeader, bool bHidden);
	int		GetHtmlText(CString& strHtml, bool bAll, bool bHeader, bool bHidden);

	void	GetColumnTitle(CString& str, int nColumn) const;
	CString	GetColumnTitle(int nColumn) const;
	int		GetColumnFormat(int nColumn) const;

	inline bool HasHeader() const { return 0 == (GetStyle() & LVS_NOCOLUMNHEADER); }

	inline TCHAR GetTextSeparationChar() const { return m_cTextSep; }
	inline TCHAR GetCsvQuoteChar()  const { return m_cCsvQuote; }
	inline TCHAR GetCsvSeparationChar() const { return m_cCsvSep; }
	inline int GetHtmlBorderWidth() const { return m_nHtmlBorder; }

	inline void	SetDragHighLight(bool b) { m_bDragHighLight = b; }
	inline void	SetDragImageType(int n, bool b)	{ m_nDragImageType = n; m_bAllowDropDesc = b; }
	inline void SetDragImageScale(int n) { m_nDragImageScale = n; }
	inline void SetDragRendering(bool b) { m_bDelayRendering = b; }
	inline void SetTextSeparationChar(TCHAR c) { m_cTextSep = c; }
	inline void SetCsvQuoteChar(TCHAR c) { ASSERT(c); m_cCsvQuote = c; }
	inline void SetCsvSeparationChar(TCHAR c) { ASSERT(c); m_cCsvSep = c; }
	inline void SetHtmlBorderWidth(int n) { ASSERT(n >= 0 && n < 10); m_nHtmlBorder = n; }

protected:
	// OLE Drag&Drop target support functions.
	static DROPEFFECT CALLBACK OnDragEnter(CWnd *pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	static DROPEFFECT CALLBACK OnDragOver(CWnd *pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	static DROPEFFECT CALLBACK OnDropEx(CWnd *pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point);
	static void CALLBACK OnDragLeave(CWnd *pWnd);
	static void CALLBACK OnDoScroll(CWnd *pWnd, int nVert, int nHorz);
	static BOOL CALLBACK OnRenderData(CWnd *pWnd, const COleDataSourceEx *pDataSrc, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium);

	virtual DROPEFFECT OnDropEx(COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point);

	COleDataSourceEx* CacheOleData(bool bClipboard, bool bAll = false, bool bHeader = false, bool bHidden = false);
	void CacheVirtualFiles(COleDataSourceEx *pDataSrc, const CString& strPlain, const CString& strCsv, const CString& strHtml);
	COleDataSourceEx* DelayRenderOleData(bool bVirtualFiles);
	void	RedrawOldDragItem();
	DROPEFFECT GetDropEffect(DWORD dwKeyState, CPoint point);
	
	void	DrawBitmapText(CDC& memDC, const CString& str, const CRect& rc, int nAlign) const;
	void	DrawBitmapImage(CDC& memDC, CImageList *pImageList, int nImage, CRect& rc, const CSize& szImage) const;

	int		InitGetTextHelper(CGetTextHelper& Options, bool bAll, bool bHeader, bool bHidden);
	int		GetPlainTextLength(const CGetTextHelper& Options) const;
	int		GetPlainText(CString& str, const CGetTextHelper& Options) const;
	void	GetTextLine(CString& strLine, int nItem, const CGetTextHelper& Options) const;
	int		GetCsvText(CString& strCsv, const CGetTextHelper& Options) const;
	void	GetCsvLine(CString& strLine, int nItem, const CGetTextHelper& Options) const;
	int		GetHtmlText(CString& strHtml, const CGetTextHelper& Options);
	LPCTSTR TextToHtml(CString& str) const;
	void	GetHtmlLine(CString& strLine, int nItem, const CGetTextHelper& Options) const;
	void	GetCellText(CString& str, int nItem, int nColumn, const CGetTextHelper& Options) const;

	bool		m_bDragHighLight;
	bool		m_bDelayRendering;
	bool		m_bCanDrop;
	bool		m_bDragging;
	bool		m_bPadText;
	bool		m_bAllowDropDesc;					// allow drop description text (Vista and later)
	bool		m_bCsvQuoteSpace;
	bool		m_bHtmlFont;
	int			m_nDragImageType;
	int			m_nDragImageScale;
	int			m_nDragItem;
	int			m_nDragSubItem;
	TCHAR		m_cTextSep;
	TCHAR		m_cCsvQuote;
	TCHAR		m_cCsvSep;
	int			m_nHtmlBorder;
	COleDropTargetEx	m_DropTarget;				// OLE Drag&Drop target support

protected:

//	virtual void PreSubclassWindow();

	afx_msg void OnNMCustomdraw(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnBegindrag(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
public:
	// Context menu and accelerator handlers
	afx_msg void OnEditCopy();
	afx_msg void OnEditCopyAll();
	afx_msg void OnEditSelectAll();
	afx_msg void OnEditDeselect();

	DECLARE_MESSAGE_MAP()
};

