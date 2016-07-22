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

#include <ShlObj.h>			// Shell data objects
#include <ShObjIdl.h>		// IDragSourceHelper2 interface

// Common clipboard text formats using _STR postfix to avoid collision with definitions from other files
//  using CF_xxx and CFSTR_xxx.
// Note that the format names are not case-sensitive.
// See richedit.h for RTF related format names.
// See shlobj.h for shell related format names.
// See MFC source oleinit.cpp for OLE related format names.
// See urlmon.h for MIME format names used by IE.
#define CF_CSV_STR					_T("CSV")
#define CF_HTML_STR					_T("HTML Format")
#define CF_RTF_STR					_T("Rich Text Format")

// Definitions that are not present when supporting old Windows versions,
//  or using old SDK header files.

// Vista clipboard format names (shlobj.h)
#ifndef CFSTR_FILE_ATTRIBUTES_ARRAY
#define CFSTR_FILE_ATTRIBUTES_ARRAY		_T("File Attributes Array")
#endif
#ifndef CFSTR_INVOKECOMMAND_DROPPARAM
#define CFSTR_INVOKECOMMAND_DROPPARAM	_T("InvokeCommand DropParam")
#endif
#ifndef CFSTR_SHELLDROPHANDLER
#define CFSTR_SHELLDROPHANDLER			_T("DropHandlerCLSID")
#endif
#ifndef CFSTR_DROPDESCRIPTION
#define CFSTR_DROPDESCRIPTION			_T("DropDescription")
#endif

// Vista: DROPDESCRIPTION and FILE_ATTRIBUTES_ARRAY (ShlObj.h)
#if !defined(NTDDI_VERSION) || (NTDDI_VERSION < NTDDI_VISTA)

// Vista: FILE_ATTRIBUTES_ARRAY struct
typedef struct
{
    UINT cItems;                    // number of items in rgdwFileAttributes array
    DWORD dwSumFileAttributes;      // all of the attributes ORed together
    DWORD dwProductFileAttributes;  // all of the attributes ANDed together
    DWORD rgdwFileAttributes[1];    // array
} FILE_ATTRIBUTES_ARRAY;            // clipboard format definition for CFSTR_FILE_ATTRIBUTES_ARRAY

// Vista: DROPDESCRIPTION type member
// NOTE: DROPIMAGE_NOIMAGE has been added with Windows 7
typedef enum  { 
  DROPIMAGE_INVALID  = -1,
  DROPIMAGE_NONE     = 0,
  DROPIMAGE_COPY     = DROPEFFECT_COPY,
  DROPIMAGE_MOVE     = DROPEFFECT_MOVE,
  DROPIMAGE_LINK     = DROPEFFECT_LINK,
  DROPIMAGE_LABEL    = 6,
  DROPIMAGE_WARNING  = 7,
  DROPIMAGE_NOIMAGE  = 8
} DROPIMAGETYPE;

// Vista: DROPDESCRIPTION structure
typedef struct _DROPDESCRIPTION {
  DROPIMAGETYPE type;
  WCHAR         szMessage[MAX_PATH];
  WCHAR         szInsert[MAX_PATH];
} DROPDESCRIPTION;

#endif

// Vista: IDragSourceHelper2 interface (ShObjIdl.h)
// Interface definition for older SDK / Visual Studio versions
//#if !defined(NTDDI_VERSION) || (NTDDI_VERSION < NTDDI_VISTA)
#ifndef __IDragSourceHelper2_INTERFACE_DEFINED__

typedef /* [v1_enum] */ 
enum DSH_FLAGS
    {	DSH_ALLOWDROPDESCRIPTIONTEXT	= 0x1
    } 	DSH_FLAGS;

#undef INTERFACE
#define INTERFACE IDragSourceHelper2

DECLARE_INTERFACE_( IDragSourceHelper2, IDragSourceHelper )
{
    // IUnknown methods
    STDMETHOD (QueryInterface)(THIS_ REFIID riid, void **ppv) PURE;
    STDMETHOD_(ULONG, AddRef) ( THIS ) PURE;
    STDMETHOD_(ULONG, Release) ( THIS ) PURE;

    // IDragSourceHelper
    STDMETHOD (InitializeFromBitmap)(THIS_ LPSHDRAGIMAGE pshdi,
                                     IDataObject* pDataObject) PURE;
    STDMETHOD (InitializeFromWindow)(THIS_ HWND hwnd, POINT* ppt,
                                     IDataObject* pDataObject) PURE;
    // IDragSourceHelper2
    STDMETHOD (SetFlags)(THIS_ DWORD) PURE;
};

static const IID IID_IDragSourceHelper2 = 
	{ 0x83E07D0D, 0x0C5F, 0x4163, 0xBF, 0x1A, 0x60, 0xB2, 0x74, 0x05, 0x1E, 0x40 };

#endif

class CDragDropHelper
{
public:
	static inline CLIPFORMAT RegisterFormat(LPCTSTR lpszFormat)
	{ return static_cast<CLIPFORMAT>(::RegisterClipboardFormat(lpszFormat)); }

	static unsigned PopCount(unsigned n);
	static bool		IsRegisteredFormat(CLIPFORMAT cfFormat, LPCTSTR lpszFormat, bool bRegister = false);
	static HGLOBAL	CopyGlobalMemory(HGLOBAL hDst, HGLOBAL hSrc, size_t nSize);
	static bool		GetGlobalData(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat, FORMATETC& FormatEtc, STGMEDIUM& StgMedium);
	static DWORD	GetGlobalDataDWord(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat);
	static bool		SetGlobalDataDWord(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat, DWORD dwData, bool bForce = true);
	static DROPIMAGETYPE DropEffectToDropImage(DROPEFFECT dwEffect);
	static bool		ClearDescription(DROPDESCRIPTION *pDropDescription);
	static size_t	GetWideLength(LPCWSTR lpszText, size_t nSize);
	static size_t	GetMultiByteLength(LPCSTR lpszText, size_t nSize);
	static LPWSTR	MultiByteToWideChar(LPCSTR lpszAnsi, int nLen, UINT nCP = CP_THREAD_ACP);
	static LPSTR	WideCharToMultiByte(LPCWSTR lpszWide, int nLen, UINT nCP = CP_THREAD_ACP);
	static UINT		GetCodePage(LCID nLCID, bool bOem);
#ifndef _UNICODE
	static bool		CompareCodePages(UINT nCP1, UINT nCP2);
	static LPSTR	MultiByteToMultiByte(LPCSTR lpszAnsi, int nLen, UINT nCP, UINT nOutCP = CP_THREAD_ACP);
#endif
};

// A class to store user defined text descriptions.
// Used by COleDataSourceEx / COleDropSourceEx and COleDropTargetEx.
class CDropDescription 
{
public:
	CDropDescription();
	virtual ~CDropDescription();

	// Max. image type stored by this class.
	static const DROPIMAGETYPE nMaxDropImage = DROPIMAGE_NOIMAGE;

	inline bool IsValidType(DROPIMAGETYPE nType) const
	{ return (nType >= DROPIMAGE_NONE && nType <= nMaxDropImage); }

	bool	SetInsert(LPCWSTR lpszInsert);
	bool	SetText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1);
	LPCWSTR	GetText(DROPIMAGETYPE nType, bool b1) const;
	bool	SetDescription(DROPDESCRIPTION *pDropDescription, DROPIMAGETYPE nType) const;
	bool	SetDescription(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszMsg) const;
	bool	CopyText(DROPDESCRIPTION *pDropDescription, DROPIMAGETYPE nType) const;
	bool	CopyMessage(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszMsg) const;
	bool	CopyInsert(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszInsert) const;

	inline LPCWSTR GetInsert() const { return m_lpszInsert; } 
	inline bool	CopyInsert(DROPDESCRIPTION *pDropDescription) const 
	{ return CopyInsert(pDropDescription, m_lpszInsert); }

	// Check if text has been stored for image type.
	inline bool HasText(DROPIMAGETYPE nType) const
	{ return IsValidType(nType) && NULL != m_pDescriptions[nType]; }

	// Check if text for image type is empty.
	inline bool IsTextEmpty(DROPIMAGETYPE nType) const
	{ return !HasText(nType) || L'\0' == m_pDescriptions[nType][0]; }


	inline bool HasInsert() const { return NULL != m_lpszInsert; }
	inline bool IsInsertEmpty() const { return !HasInsert() || L'\0' == *m_lpszInsert; }

	inline bool HasInsert(const DROPDESCRIPTION *pDropDescription) const
	{ return pDropDescription && pDropDescription->szInsert[0] != L'\0'; }

	inline bool	ClearDescription(DROPDESCRIPTION *pDropDescription) const
	{ return CDragDropHelper::ClearDescription(pDropDescription); }

protected:
	bool CopyDescription(LPWSTR lpszDest, size_t nDstSize, LPCWSTR lpszSrc) const;

	LPWSTR m_lpszInsert;
	LPWSTR m_pDescriptions[nMaxDropImage + 1];
	LPWSTR m_pDescriptions1[nMaxDropImage + 1];
};

// Helper class to create and decode HTML clipboard format data.
class CHtmlFormat
{
public:
	CHtmlFormat();

	HGLOBAL	SetHtml(LPCTSTR lpszHtml, CWnd *pWnd = NULL);
	LPSTR	GetUtf8(LPCSTR lpData, int nSize, bool bAll);
	LPTSTR	GetHtml(LPCSTR lpData, int nSize, bool bAll);

	inline size_t GetSize() const { return m_nSize; }
	inline LPCSTR GetVersion() const { return m_strVersion.GetString(); }
	inline LPCSTR GetSourceURL() const { return m_strSrcURL.GetString(); }
	inline INT_PTR GetFragmentCount() const { return m_arStartFrag.GetCount(); }
	inline int GetStartHtml() const { return m_nStartHtml; }
	inline int GetEndHtml() const { return m_nEndHtml; }

	static LPCTSTR GetMappedFontName(const LOGFONT& lf);
	static void GetFontStyle(CString& strStyle, CWnd *pWnd);

protected:
	int		FindNoCase(LPCTSTR lpszBuffer, LPCTSTR lpszKeyword) const;
	int		FindNoCase(const CString& str, LPCTSTR lpszKeyword, int nStart = 0) const;
	LPCSTR	FindHtmlDescriptor(LPCSTR lpszBuffer, LPCSTR lpszKeyword) const;
	LPCSTR	GetHtmlDescriptorValue(LPCSTR lpszBuffer, LPCSTR lpszKeyword, int& nValue) const;
	LPCSTR	GetHtmlDescriptorValue(LPCSTR lpszBuffer, LPCSTR lpszKeyword, CStringA& strValue) const;

	int			m_nStartHtml;
	int			m_nEndHtml;
	size_t		m_nSize;
	CDWordArray	m_arStartFrag;
	CDWordArray	m_arEndFrag;
	LPCSTR		m_lpszStartHtml;
	CStringA	m_strVersion;
	CStringA	m_strSrcURL;

};

class CTextToRTF
{
public:
	CTextToRTF();

	void StringToRtf(CString& strRtf, LPCTSTR lpszText);
	LPCTSTR TextToRtf(CString& strRtf, LPCTSTR lpszText, LPCTSTR lpszFontName, int nFontSize);
	LPCTSTR TextToRtf(CString& strRtf, LPCTSTR lpszText, CWnd *pWnd);
	LPCTSTR TextToRtf(CString& strRtf, LPCTSTR lpszText, CFont *pFont, HDC hDC);

protected:
	LPCTSTR CharToRtf(wchar_t cWide);
#ifndef _UNICODE
	int ToWideChar(LPCSTR lpszText, wchar_t& cWide);
#endif

	bool m_bDelim;									// control word delimiter flag
	TCHAR m_lpszFormat[16];							// buffer for ANSI and Unicode control words

    CString* pChecker = nullptr;
};
