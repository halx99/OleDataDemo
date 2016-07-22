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

#include "OleDropSourceEx.h"

// When defined, serialized data from objects are supported.
//#define OBJ_SUPPORT

class COleDataSourceEx;

typedef BOOL (CALLBACK* OnRenderDataFunc)(CWnd*, const COleDataSourceEx*, LPFORMATETC, LPSTGMEDIUM);

// Class to generate virtual file data objects
class VirtualFileData
{
public:
	VirtualFileData()
	{ 
		m_nSize = m_nBOM = 0; 
		m_lpszFileName = NULL; 
		m_lpData = m_lpBOM = NULL; 
	}

	inline void SetFileName(LPCTSTR lpszFileName) { m_lpszFileName = lpszFileName; }
	inline void SetString(LPCSTR lpszString) { m_lpData = lpszString; m_nSize = strlen(lpszString); }
	inline void SetSize(size_t nSize) { m_nSize = nSize; }
	inline void SetData(LPCVOID lpData, size_t nSize) { m_lpData = lpData; m_nSize = nSize; }
	inline void SetBOM(LPCSTR lpszBOM) { m_lpBOM = lpszBOM; m_nBOM = strlen(lpszBOM); }
	inline void SetBOM(LPCVOID lpBOM, size_t nSize) { m_lpBOM = lpBOM; m_nBOM = nSize; }

	inline size_t GetDataSize() const { return m_nSize; }
	inline size_t GetBOMSize() const { return m_nBOM; }
	inline size_t GetTotalSize() const { return m_nSize + m_nBOM; }
	inline LPCVOID GetData() const { return m_lpData; }
	inline LPCVOID GetBOM() const { return m_lpBOM; }
	inline LPCTSTR GetFileName() const { return m_lpszFileName; }

#ifdef _DEBUG
	bool IsValid() const
	{ return m_lpszFileName && *m_lpszFileName && _tcslen(m_lpszFileName) < MAX_PATH; }
	bool IsDataValid() const
	{ return IsValid() && m_nSize && m_lpData && (m_lpBOM || 0 == m_nBOM); }
#endif
protected:
	size_t		m_nSize;							// size of m_lpBuf
	size_t		m_nBOM;								// size of m_lpBOM
	LPCTSTR		m_lpszFileName;						// file name
	LPCVOID		m_lpData;							// data
	LPCVOID		m_lpBOM;							// optional BOM
};


class COleDataSourceEx : public COleDataSource
{
public:
	COleDataSourceEx(CWnd *pWnd = NULL, OnRenderDataFunc pOnRenderData = NULL);
	virtual ~COleDataSourceEx();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif
	
	inline bool WasDragStarted() const { return m_nDragResult & DRAG_RES_STARTED; }
	inline int GetDragResult() const { return m_nDragResult; }

	DROPEFFECT DoDragDropEx(DROPEFFECT dwEffect, LPCRECT lpRectStartDrag = NULL);

	// Drop description functions
	bool AllowDropDescriptionText();
	bool SetDropDescriptionText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1, LPCWSTR lpszInsert = NULL);

	// Drag image functions
	bool SetDragImageWindow(HWND hWnd, POINT* pPoint);
	bool SetDragImage(HBITMAP hBitmap, const CPoint* pPoint, COLORREF clr);
	// Set drag image from CBitmap. NOTE: This will use the HBITMAP from the CBitmap and detach!
	inline bool SetDragImage(CBitmap& Bitmap, const CPoint* pPoint, COLORREF clr)
	{ return SetDragImage(static_cast<HBITMAP>(Bitmap.Detach()), pPoint, clr); }

	bool InitDragImage(int nResBM, const CPoint* pPoint, COLORREF clr);
	bool InitDragImage(LPCTSTR lpszExt, int nScale = 0, const CPoint* pPoint = NULL, UINT nFlags = 0);
	bool InitDragImage(CWnd* pWnd, LPCTSTR lpszText, const CPoint *pPoint = NULL, bool bHiLightColor = false);
	bool InitDragImage(CWnd* pWnd, LPCRECT lpRect, int nScale, const CPoint *pPoint = NULL, COLORREF clr = CLR_INVALID);
	bool InitDragImage(HBITMAP hBitmap, int nScale, const CPoint *pPoint = NULL, COLORREF clrBk = CLR_INVALID);

	// Data caching functions
	bool CacheLocale(LCID nLCID = 0);
#ifndef _UNICODE
	bool CacheUnicode(CLIPFORMAT cfFormat, LPCSTR lpszAnsi, DWORD dwTymed = TYMED_HGLOBAL);
#endif
	bool CacheMultiByte(LPCTSTR lpszFormat, LPCWSTR lpszData, UINT nCP = CP_THREAD_ACP, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheMultiByte(CLIPFORMAT cfFormat, LPCWSTR lpszData, UINT nCP = CP_THREAD_ACP, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheText(LPCTSTR lpszText, bool bClipboard, bool bOem = false, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheRtfFromText(LPCTSTR lpszText, CWnd *pWnd = NULL, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheHtml(LPCTSTR lpszHtml, CWnd *pWnd = NULL, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheVirtualFiles(unsigned nFiles, const VirtualFileData *pData, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheFileDescriptor(unsigned nFiles, const VirtualFileData *pData);
	bool CacheBitmap(HBITMAP hBitmap, bool bClipboard, HPALETTE hPal = NULL, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheDib(HBITMAP hBitmap, bool bV5, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheBitmapFromText(LPCTSTR lpszText, CWnd *pWnd);
	bool CacheBitmapAsFile(HBITMAP hBitmap, LPCTSTR lpszFileName, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheDdb(HBITMAP hBitmap, HPALETTE hPal, bool bCopyObjects);
	bool CacheImage(HBITMAP hBitmap, LPCTSTR lpszFormat, DWORD dwTymed = TYMED_HGLOBAL);
	bool CacheImage(HBITMAP hBitmap, CLIPFORMAT cfFormat, DWORD dwTymed = TYMED_HGLOBAL);
	inline bool CacheTiff(HBITMAP hBitmap, DWORD dwTymed = TYMED_HGLOBAL) 
	{ return CacheImage(hBitmap, CF_TIFF, dwTymed); }
	inline bool CacheMimeBmp(HBITMAP hBitmap, DWORD dwTymed = TYMED_HGLOBAL) 
	{ return CacheImage(hBitmap, _T("image/bmp"), dwTymed); }
	inline bool CacheMimePng(HBITMAP hBitmap, DWORD dwTymed = TYMED_HGLOBAL) 
	{ return CacheImage(hBitmap, _T("image/png"), dwTymed); }
	inline bool CacheMimeJpeg(HBITMAP hBitmap, DWORD dwTymed = TYMED_HGLOBAL) 
	{ return CacheImage(hBitmap, _T("image/jpeg"), dwTymed); }
	inline bool CacheMimeGif(HBITMAP hBitmap, DWORD dwTymed = TYMED_HGLOBAL) 
	{ return CacheImage(hBitmap, _T("image/gif"), dwTymed); }
#ifdef OBJ_SUPPORT
	bool CacheObjectData(CObject& Obj, LPCTSTR lpszFormat);
#endif
	bool CacheSingleFileAsHdrop(LPCTSTR lpszFileName);

#ifdef UNICODE
	inline bool CacheRtf(LPCTSTR lpszRtf, DWORD dwTymed = TYMED_HGLOBAL)
	{ return CacheMultiByte(CF_RTF_STR, lpszRtf, 20127, dwTymed); }	// 20127 is US-ASCII
	inline bool CacheCsv(LPCTSTR lpszCsv, DWORD dwTymed = TYMED_HGLOBAL)
	{ return CacheMultiByte(CF_CSV_STR, lpszCsv, CP_THREAD_ACP, dwTymed); }
#else
	inline bool CacheRtf(LPCTSTR lpszRtf, DWORD dwTymed = TYMED_HGLOBAL)
	{ return CacheString(CF_RTF_STR, lpszRtf, dwTymed); }
	inline bool CacheCsv(LPCTSTR lpszCsv, DWORD dwTymed = TYMED_HGLOBAL)
	{ return CacheString(CF_CSV_STR, lpszCsv, dwTymed); }
#endif

	// Render data functions

	virtual BOOL OnRenderData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium);

	bool RenderText(CLIPFORMAT cfFormat, LPCTSTR lpszText, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const;
	bool RenderUnicode(LPCTSTR lpszText, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const;
	bool RenderMultiByte(LPCTSTR lpszText, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, UINT nCP = CP_THREAD_ACP) const;
	bool RenderHtml(LPCTSTR lpszHtml, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, CWnd *pWnd) const;
	bool RenderBitmap(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const;
	bool RenderDdb(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const;
	bool RenderDib(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const;
	size_t RenderImage(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const;
	size_t RenderImage(HBITMAP hBitmap, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, const GUID& FileType) const;
	// Render string as OEM text.
//	inline bool RenderOEM(LPCTSTR lpszText, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const
//	{ return RenderMultiByte(lpszText, lpFormatEtc, lpStgMedium, GetOEMCodePage()); }

	bool BitmapFromText(CBitmap& Bitmap, CWnd* pWnd, LPCTSTR lpszText) const;

	static BOOL BitmapFromWnd(CWnd *pWnd, CBitmap &Bitmap, CPalette *pPal, LPCRECT lpRect = NULL, COLORREF *lpBkClr = NULL, int nScale = 0);
	static BOOL BitmapFromWnd(CWnd *pWnd, CBitmap &Bitmap);
	static BOOL PrintWndToBitmap(CWnd *pWnd, CBitmap &Bitmap, bool bClient);

protected:
	bool				m_bSetDescriptionText;		// Description text string(s) has been specified
	int					m_nDragResult;				// DoDragDropEx result flags
	IDragSourceHelper*	m_pDragSourceHelper;		// Drag image helper
	IDragSourceHelper2*	m_pDragSourceHelper2;		// Drag image helper 2 (SetFlags function)
	CDropDescription	m_DropDescription;			// Optional text descriptions
	CWnd				*m_pWnd;					// Associated window
	OnRenderDataFunc	m_pOnRenderData;			// RenderData callback function

	CString		CreateTempFileName(LPCTSTR lpszExt = NULL) const;
	LPOLESTR	CreateOleString(LPCTSTR lpszStr) const;
	DWORD		GetLowestColorBit(DWORD dwBitField) const;
	HBITMAP		ReplaceBlack(HBITMAP hBitmap) const;
	HGLOBAL		CreateDib(HBITMAP hBitmap, size_t& nSize, bool bV5, bool bFile) const;
	bool		CacheString(LPCTSTR lpszFormat, LPCTSTR lpszText, DWORD dwTymed = TYMED_HGLOBAL);
	bool		CacheString(CLIPFORMAT cfFormat, LPCTSTR lpszText, DWORD dwTymed = TYMED_HGLOBAL);
	bool		CacheFromBuffer(CLIPFORMAT cfFormat, LPCVOID lpData, size_t nSize, DWORD dwTymed, LPFORMATETC lpFormatEtc = NULL);
	bool		CacheFromGlobal(CLIPFORMAT cfFormat, HGLOBAL hGlobal, size_t nSize, DWORD dwTymed, LPFORMATETC lpFormatEtc = NULL);
	bool		RenderFromBuffer(LPCVOID lpData, size_t nSize, LPCVOID lpPrefix, size_t nPrefixSize, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const;
	bool		RenderFromGlobal(HGLOBAL hGlobal, size_t nSize, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium) const;
	UINT		GetDibColorTable(HBITMAP hBitmap, UINT nColors, RGBQUAD* pColors) const;

	static int	ScaleSize(int nScale, CSize& size, HDC hDC = NULL);

public:
	// Drag source helper support.
	// See http://www.codeproject.com/Articles/3530/DragSourceHelper-MFC
	// Helper methods to fix IDropSourceHelper.
	// Must implement all pure virtual functions.
	BEGIN_INTERFACE_PART(DataObjectEx, IDataObject)
		INIT_INTERFACE_PART(COleDataSourceEx, DataObjectEx)
		// IDataObject
		STDMETHOD(GetData)(LPFORMATETC, LPSTGMEDIUM);
		STDMETHOD(GetDataHere)(LPFORMATETC, LPSTGMEDIUM);
		STDMETHOD(QueryGetData)(LPFORMATETC);
		STDMETHOD(GetCanonicalFormatEtc)(LPFORMATETC, LPFORMATETC);
		STDMETHOD(SetData)(LPFORMATETC, LPSTGMEDIUM, BOOL);
		STDMETHOD(EnumFormatEtc)(DWORD, LPENUMFORMATETC*);
		STDMETHOD(DAdvise)(LPFORMATETC, DWORD, LPADVISESINK, LPDWORD);
		STDMETHOD(DUnadvise)(DWORD);
		STDMETHOD(EnumDAdvise)(LPENUMSTATDATA*);
	END_INTERFACE_PART(DataObjectEx)

	DECLARE_INTERFACE_MAP()
};
