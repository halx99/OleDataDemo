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

// CMyStatic

class CMyStatic : public CStatic
{
	DECLARE_DYNAMIC(CMyStatic)

public:
	CMyStatic();
	virtual ~CMyStatic();

	enum DropType
	{
		DropNothing,
		DropBitmap,
		DropImageFile,
		DropDragImage
	};

	void Init(int nResID = 0);
	bool ReplaceBitmap(HBITMAP hBitmap);
	bool LoadBitmap(int nID);
	bool LoadFromFile(LPCTSTR lpszFileName);
	bool CanPaste() const;
	inline bool CanCopy() const { return NULL != GetBitmap(); }
	inline void SetDragImageType(int n, bool b) { m_nDragImageType = n; m_bAllowDropDesc = b; }
	inline void SetDragImageScale(int n) { m_nDragImageScale = n; }
	inline void SetDragRendering(bool b) { m_bDelayRendering = b; }

protected:
	bool				m_bAllowDropDesc;
	bool				m_bDelayRendering;
	bool				m_bDragging;				// Internal var; set when this is drag & drop source
	int					m_nDragImageType;
	int					m_nDragImageScale;
	DropType			m_nDropType;
	COleDropTargetEx	m_DropTarget;				// OLE Drag&Drop target support

protected:
//	virtual void PreSubclassWindow();

	// OLE Drag&Drop target handlers.
	LRESULT	OnDragEnter(WPARAM wParam, LPARAM lParam);
	LRESULT	OnDragOver(WPARAM wParam, LPARAM lParam);
	LRESULT	OnDropEx(WPARAM wParam, LPARAM lParam);

	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);

	bool IsSupportedFileExtension(LPCTSTR lpszFileName) const;
	bool CanLoadFromFile(COleDataObject *pDataObject) const;
	DropType GetDropType(COleDataObject *pDataObject, bool bDropDragImage = false) const;
	bool PutOleData(COleDataObject *pDataObject, DropType nType);
	COleDataSourceEx *CacheOleData(HBITMAP hBitmap, bool bClipboard);
	COleDataSourceEx *DelayRenderOleData(HBITMAP hBitmap);

	static BOOL CALLBACK OnRenderData(CWnd *pWnd, const COleDataSourceEx *pDataSrc, LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium);

public:
	afx_msg void OnEditPaste();
	afx_msg void OnEditCopy();
	DECLARE_MESSAGE_MAP()
};


