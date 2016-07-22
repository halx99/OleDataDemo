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

class COleDataDemoDlg;

// CInfoList

class CDataInfo
{
public:
	CDataInfo(const COleDropTargetEx *pDropTarget, bool bCurrentLocale = false);
	~CDataInfo();

	unsigned GetInfo(DWORD tymed, const FORMATETC& FormatEtcIn, COleDataObject* pDataObj);
	CString GetTymedString(DWORD dwTymed) const;
	CString GetDeviceString() const;
	LPCTSTR GetAspectString() const;
	LPCTSTR PrintSize(size_t nSize) const;
	CString GetDropEffectString(DROPEFFECT dwEffect) const;
	LPCTSTR GetFixedFormatName(CLIPFORMAT cfFormat) const;

	inline size_t GetSize() const { return m_nSize; }
	inline LPCTSTR GetFormatName() const { return m_strFormat.GetString(); }
	inline LPCTSTR GetDataType() const { return m_strDataType.GetString(); }
	inline LPCTSTR GetContent() const { return m_strContent.GetString(); }
	inline bool HasDataType() const { return !m_strDataType.IsEmpty(); }
	inline bool HasContent() const { return !m_strContent.IsEmpty(); }
	inline CLIPFORMAT GetFormat() const { return m_FormatEtc.cfFormat; }
	inline DWORD GetTymed() const { return m_FormatEtc.tymed; }
	inline DWORD GetUsedTymed() const { return m_nUsedTymed; }
	inline LONG GetIndex() const { return m_FormatEtc.lindex; }
	inline DVTARGETDEVICE *GetTargetDevice() const { return m_FormatEtc.ptd; }
	inline DWORD GetAspect() const { return m_FormatEtc.dwAspect; }
	inline IUnknown* GetUnkForRelease() const { return m_pUnkForRelease; }

protected:
	LPVOID AllocBuffer(size_t nSize);
	void FreeBuffer();
	void GetStandardFormatInfo(COleDataObject* pDataObj, STGMEDIUM& m_stgm);
	void GetCustomFormatInfo(COleDataObject* pDataObj, STGMEDIUM& m_stgm);
	void GetBinaryContent();
	void GetTextInfo();
	void GetBitmapInfo(const BITMAPINFOHEADER *lpInfo);
	void GetObjectDescriptor();
	void GetLinkSource(STGMEDIUM& m_stgm);
	void GetHyperLink(COleDataObject* pDataObj, STGMEDIUM& m_stgm);
	unsigned GetDropFileList(LPCTSTR lpszType);
	unsigned GetStringList(UINT nCP, LPCTSTR lpszType);
	unsigned ParseStringList(const BYTE *p, LPCTSTR lpszType);
	void GetFileContentsInfo(COleDataObject *pDataObject);
	int CheckForString();
	bool CheckBinaryFormats();
	void AppendCLSID(const CLSID& id, LPCTSTR lpszPrefix);
	void AppendOleStr(LPOLESTR lpOleStr, LPCTSTR lpszPrefix);
	void SeparateContent();
	void AppendString(CString& str, LPCTSTR lpszText, LPCTSTR lpszSep = _T(" | ")) const;
	bool SetFormatName(CLIPFORMAT cfFormat);

	// Check if specific number of data bytes can be accessed 
	inline bool HasData(size_t n) const { return m_lpData && m_nBufSize >= n; }
	// Check if all data bytes can be accessed (always true with HGLOBAL data)
	inline bool HasAllData() const { return m_lpData && m_nBufSize == m_nSize; }
	inline const BYTE * GetByteData() const { return static_cast<const BYTE*>(m_lpData); }

	static const LARGE_INTEGER m_LargeNull;

	bool		m_bBufferAllocated;
	UINT		m_nCP;
	size_t		m_nSize;
	size_t		m_nBufSize;
	LPVOID		m_lpData;
	TCHAR		m_lpszThSep[5];
	FORMATETC	m_FormatEtc;
	DWORD		m_nUsedTymed;
	IUnknown	*m_pUnkForRelease;
	CString		m_strText;
	CString		m_strFormat;
	CString		m_strDataType;
	CString		m_strContent;
	const COleDropTargetEx *m_pDropTarget;
};

class CInfoList : public CListCtrl
{
	DECLARE_DYNAMIC(CInfoList)

public:
	CInfoList();
	virtual ~CInfoList();

	void Init(COleDataDemoDlg *p = NULL, bool bCurrentLocale = false);
	void UpdateFromClipboard();
	void UpdateFromData(COleDataObject *pDataObj);
	void SizeColumns(unsigned nMask = 0);
	inline void DefaultColumns() { SizeColumns(m_nDefaultColumns); }
	inline void AllColumns() { SizeColumns(UINT_MAX); }
	inline void AutoColumns() { SizeColumns(m_nAutoColumns); }

	enum InfoColumns
	{
		FormatColumn = 0,
		TymedColumn,
		UsedTymedColumn,
		DeviceColumn,
		AspectColumn,
		IndexColumn,
		ReleaseColumn,
		SizeColumn,
		TypeColumn,
		ContentColumn
	};

protected:
	bool				m_bAllTymeds;
	bool				m_bCurrentLocale;
	unsigned			m_nVisibleColumns;
	unsigned			m_nAutoColumns;
	HWND				m_hWndChain;
	COleDropTargetEx	m_DropTarget;
	COleDataDemoDlg *	m_pParentDlg;

	static const unsigned m_nDefaultColumns;

	int AddItem(const CDataInfo& Info);

	static DROPEFFECT CALLBACK OnDragEnter(CWnd *pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	static DROPEFFECT CALLBACK OnDragOver(CWnd *pWnd, COleDataObject* pDataObject, DWORD dwKeyState, CPoint point);
	static DROPEFFECT CALLBACK OnDropEx(CWnd *pWnd, COleDataObject* pDataObject, DROPEFFECT dropDefault, DROPEFFECT dropList, CPoint point);
	DROPEFFECT GetDropEffect(CPoint point);

public:
	afx_msg void OnDrawClipboard();
	afx_msg void OnChangeCbChain(HWND hWndRemove, HWND hWndAfter);
	afx_msg void OnDestroy();
	DECLARE_MESSAGE_MAP()
};


