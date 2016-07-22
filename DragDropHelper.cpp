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
#include "DragDropHelper.h"

// Counts the number of one bits (population count) 
/*static*/ unsigned CDragDropHelper::PopCount(unsigned n)
{
#if defined(__INTRIN_H_) && (_MSC_VER >= 1500)		// intrin.h included and VS 2008 or later
	// Using this requires that the CPU supports the popcnt instruction.
	return __popcnt(n);
#else
	// popcount implementation
	// See http://graphics.stanford.edu/~seander/bithacks.html, "Counting bits set"
	unsigned nPopCnt;
	for (nPopCnt = 0; n; nPopCnt++) 
		n &= n - 1;
	return nPopCnt;
#endif
}

// Check if cfFormat is the registered format specified by lpszFormat.
// bRegister may be true when it is known that the format has been registered already.
// Otherwise, it should be false to avoid registering unused formats.
/*static*/ bool CDragDropHelper::IsRegisteredFormat(CLIPFORMAT cfFormat, LPCTSTR lpszFormat, bool bRegister /*= false*/)
{
	bool bRet = false;
	if (cfFormat >= 0xC000)
	{
		if (bRegister)
			bRet = (cfFormat == RegisterFormat(lpszFormat));
		else
		{
			TCHAR lpszName[128];
			if (::GetClipboardFormatName(cfFormat, lpszName, ARRAYSIZE(lpszName)))
				bRet = (0 == _tcsicmp(lpszFormat, lpszName));
		}
	}
	return bRet;
}

// Get corresponding drop description image type from drop effect. 
// With valid code, nType = static_cast<DROPIMAGETYPE>(dwEffect & ~DROPEFFECT_SCROLL)
//  can be used. But drop handlers may return effects with multiple bits set.
/*static*/ DROPIMAGETYPE CDragDropHelper::DropEffectToDropImage(DROPEFFECT dwEffect)
{
	DROPIMAGETYPE nImageType = DROPIMAGE_INVALID;
	dwEffect &= ~DROPEFFECT_SCROLL;
	if (DROPEFFECT_NONE == dwEffect)
		nImageType = DROPIMAGE_NONE;
	else if (dwEffect & DROPEFFECT_MOVE)
		nImageType = DROPIMAGE_MOVE;
	else if (dwEffect & DROPEFFECT_COPY)
		nImageType = DROPIMAGE_COPY;
	else if (dwEffect & DROPEFFECT_LINK)
		nImageType = DROPIMAGE_LINK;
	return nImageType;
}

// Copy data from one global memory block to another.
// When hDst is NULL, it is allocated here.
/*static*/ HGLOBAL CDragDropHelper::CopyGlobalMemory(HGLOBAL hDst, HGLOBAL hSrc, size_t nSize)
{
	ASSERT(hSrc);

	if (0 == nSize)
		nSize = ::GlobalSize(hSrc);
	if (nSize)
	{
		if (NULL == hDst)
			hDst = ::GlobalAlloc(GMEM_MOVEABLE, nSize);
		else if (nSize > ::GlobalSize(hDst))
			hDst = NULL;
		if (hDst)
		{
			::CopyMemory(::GlobalLock(hDst), ::GlobalLock(hSrc), nSize);
			::GlobalUnlock(hDst);
			::GlobalUnlock(hSrc);
		}
	}
	return hDst;
}

// Helper function to access global data objects.
// Call ReleaseStgMedium(&StgMedium) when this succeeds and finished with processing.
/*static*/ bool CDragDropHelper::GetGlobalData(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat, FORMATETC& FormatEtc, STGMEDIUM& StgMedium)
{
	ASSERT(pIDataObj);
	ASSERT(lpszFormat && *lpszFormat);

	bool bRet = false;
	FormatEtc.cfFormat = RegisterFormat(lpszFormat);
	FormatEtc.ptd = NULL;
	FormatEtc.dwAspect = DVASPECT_CONTENT;
	FormatEtc.lindex = -1;
	FormatEtc.tymed = TYMED_HGLOBAL;
	if (SUCCEEDED(pIDataObj->QueryGetData(&FormatEtc)))
	{
		if (SUCCEEDED(pIDataObj->GetData(&FormatEtc, &StgMedium)))
		{
			bRet = (TYMED_HGLOBAL == StgMedium.tymed);
			if (!bRet)
				::ReleaseStgMedium(&StgMedium);
		}
	}
	return bRet;
}

// Helper function to read global data DWORD objects.
/*static*/ DWORD CDragDropHelper::GetGlobalDataDWord(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat)
{
	DWORD dwData = 0;
	FORMATETC FormatEtc;
	STGMEDIUM StgMedium;
	if (GetGlobalData(pIDataObj, lpszFormat, FormatEtc, StgMedium))
	{
		ASSERT(::GlobalSize(StgMedium.hGlobal) >= sizeof(DWORD));
		dwData = *(static_cast<LPDWORD>(::GlobalLock(StgMedium.hGlobal)));
		::GlobalUnlock(StgMedium.hGlobal);
		::ReleaseStgMedium(&StgMedium);
	}
	return dwData;
}

// Create global DWORD data object or set value for existing objects.
// When bForce is true, non-existing objects are created even if dwData is zero.
/*static*/ bool CDragDropHelper::SetGlobalDataDWord(LPDATAOBJECT pIDataObj, LPCTSTR lpszFormat, DWORD dwData, bool bForce /*= true*/)
{
	bool bSet = false;
	FORMATETC FormatEtc;
	STGMEDIUM StgMedium;
	// Check if object exists already
	if (GetGlobalData(pIDataObj, lpszFormat, FormatEtc, StgMedium))
	{
		LPDWORD pData = static_cast<LPDWORD>(::GlobalLock(StgMedium.hGlobal));
		bSet = (*pData != dwData);
		*pData = dwData;
		::GlobalUnlock(StgMedium.hGlobal);
		if (bSet)
			bSet = SUCCEEDED(pIDataObj->SetData(&FormatEtc, &StgMedium, TRUE));
		if (!bSet)									// not changed or setting failed
			::ReleaseStgMedium(&StgMedium);
	}
	else if (dwData || bForce)
	{
		StgMedium.hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD));
		if (StgMedium.hGlobal)
		{
			LPDWORD pData = static_cast<LPDWORD>(::GlobalLock(StgMedium.hGlobal));
			*pData = dwData;
			::GlobalUnlock(StgMedium.hGlobal);
			StgMedium.tymed = TYMED_HGLOBAL;
			StgMedium.pUnkForRelease = NULL;
			bSet = SUCCEEDED(pIDataObj->SetData(&FormatEtc, &StgMedium, TRUE));
			if (!bSet)
				::GlobalFree(StgMedium.hGlobal);
		}
	}
	return bSet;
}

// Clear drop description data.
// Returns true when data has been changed.
/*static*/ bool CDragDropHelper::ClearDescription(DROPDESCRIPTION *pDropDescription)
{
	ASSERT(pDropDescription);

	bool bChanged = 
		pDropDescription->type != DROPIMAGE_INVALID ||
		pDropDescription->szMessage[0] != L'\0' || 
		pDropDescription->szInsert[0] != L'\0'; 
	pDropDescription->type = DROPIMAGE_INVALID;
	pDropDescription->szMessage[0] = pDropDescription->szInsert[0] = L'\0';
	return bChanged;
}

// Get length of Unicode string.
// Used with strings from data objects.
// These strings may be not NULL terminated and the buffer size may be larger
//  than necessary.
/*static*/ size_t CDragDropHelper::GetWideLength(LPCWSTR lpszText, size_t nSize)
{
	ASSERT(lpszText);
	ASSERT(nSize);

	size_t nLen = 0;
	size_t nMaxLen = nSize / sizeof(WCHAR);
	while (lpszText[nLen] && nLen < nMaxLen)
		++nLen;
	return nLen;
}

// Get length of multi byte string.
// Used with strings from global data objects.
// These strings may be not NULL terminated and the buffer size may be larger
//  than necessary.
// NOTE: This returns number of bytes!
/*static*/ size_t CDragDropHelper::GetMultiByteLength(LPCSTR lpszText, size_t nSize)
{
	ASSERT(lpszText);
	ASSERT(nSize);

	size_t nLen = 0;
	while (lpszText[nLen] && nLen < nSize)
		++nLen;
	return nLen;
}

// Convert multi byte string to allocated Unicode string.
/*static*/ LPWSTR CDragDropHelper::MultiByteToWideChar(LPCSTR lpszAnsi, int nLen, UINT nCP /*= CP_THREAD_ACP*/)
{
	ASSERT(lpszAnsi);

	LPWSTR lpszWide = NULL;
	if (nLen < 0)
		nLen = static_cast<int>(strlen(lpszAnsi));
	int nWideLen = ::MultiByteToWideChar(nCP, 0, lpszAnsi, nLen, NULL, 0);
	if (nWideLen)
	{
		try
		{
			lpszWide = new WCHAR[nWideLen + 1];
			if (0 == ::MultiByteToWideChar(nCP, 0, lpszAnsi, nLen, lpszWide, nWideLen))
			{
				delete [] lpszWide;
				lpszWide = NULL;
			}
			else
				lpszWide[nWideLen] = L'\0';
		}
		catch (CMemoryException *pe)
		{
			pe->Delete();
		}
	}
	return lpszWide;
}

// Convert Unicode string to allocated multi byte string.
/*static*/ LPSTR CDragDropHelper::WideCharToMultiByte(LPCWSTR lpszWide, int nLen, UINT nCP /*= CP_THREAD_ACP*/)
{
	ASSERT(lpszWide);

	LPSTR lpszAnsi = NULL;
	if (nLen < 0)
		nLen = static_cast<int>(wcslen(lpszWide));
	int nAnsiLen = ::WideCharToMultiByte(nCP, 0, lpszWide, nLen, NULL, 0, NULL, NULL);
	if (nAnsiLen)
	{
		try
		{
			lpszAnsi = new char[nAnsiLen + 1];
			if (0 == ::WideCharToMultiByte(nCP, 0, lpszWide, nLen, lpszAnsi, nAnsiLen, NULL, NULL))
			{
				delete [] lpszAnsi;
				lpszAnsi = NULL;
			}
			else
				lpszAnsi[nAnsiLen] = '\0';
		}
		catch (CMemoryException *pe)
		{
			pe->Delete();
		}
	}
	return lpszAnsi;
}

// Get code page number for specified locale.
/*static*/ UINT CDragDropHelper::GetCodePage(LCID nLCID, bool bOem)
{
	UINT nCP = 0;
	TCHAR lpszCP[8];
	if (0 == LANGIDFROMLCID(nLCID))
		nLCID = ::GetThreadLocale();
	if (::GetLocaleInfo(nLCID, 
		bOem ? LOCALE_IDEFAULTCODEPAGE : LOCALE_IDEFAULTANSICODEPAGE, 
		lpszCP, ARRAYSIZE(lpszCP)))
	{
		nCP = _tstoi(lpszCP);
	}
	return nCP;
}

#ifndef _UNICODE
// Compare effective code pages.
// Returns true if code pages are identical.
/*static*/ bool CDragDropHelper::CompareCodePages(UINT nCP1, UINT nCP2)
{
	CPINFOEX CpInfo1, CpInfo2;
	VERIFY(::GetCPInfoEx(nCP1, 0, &CpInfo1));
	VERIFY(::GetCPInfoEx(nCP2, 0, &CpInfo2));
	return CpInfo1.CodePage == CpInfo2.CodePage;
}

// Convert multi byte string from one code page to another.
/*static*/ LPSTR CDragDropHelper::MultiByteToMultiByte(LPCSTR lpszAnsi, int nLen, UINT nCP, UINT nOutCP /*= CP_THREAD_ACP*/)
{
	LPSTR lpszOut = NULL;
	if (!CompareCodePages(nCP, nOutCP))
	{
		LPCWSTR lpszWide = MultiByteToWideChar(lpszAnsi, nLen, nCP);
		if (lpszWide)
		{
			lpszOut = WideCharToMultiByte(lpszWide, -1, nOutCP);
			delete [] lpszWide;
		}
	}
	else
	{
		if (nLen < 0)
			nLen = strlen(lpszAnsi);
		try
		{
			lpszOut = new char[nLen + 1];
			::CopyMemory(lpszOut, lpszAnsi, nLen);
			lpszOut[nLen] = '\0';
		}
		catch (CMemoryException& e)
		{
			e.Delete();
		}
	}
	return lpszOut;
}
#endif // #ifndef _UNICODE


//
// A class to store user defined text descriptions
//

CDropDescription::CDropDescription()
{
	m_lpszInsert = NULL;
	::ZeroMemory(m_pDescriptions, sizeof(m_pDescriptions));
	::ZeroMemory(m_pDescriptions1, sizeof(m_pDescriptions1));
}

CDropDescription::~CDropDescription()
{
	delete [] m_lpszInsert;
	for (unsigned n = 0; n <= nMaxDropImage; n++)
	{
		delete [] m_pDescriptions[n];
		delete [] m_pDescriptions1[n];
	}
}

// Set text for specified image type.
// lpszText:  String to be used without szInsert.
// lpszText1: String to be used with szInsert.
//            This may be NULL. Then lpszText is returned when requesting the string with szInsert.
// NOTES:
//  - Because the %1 placeholder is used to insert the szInsert text, a single percent character
//    inside the strings must be esacped by another one.
//  - When passing NULL for lpszText (or just not calling this for a specific image type),
//    the default system text is shown.
//    To avoid this, pass an empty string. However, this will result in an empty text area
//     shown below and right of the new style cursor.
bool CDropDescription::SetText(DROPIMAGETYPE nType, LPCWSTR lpszText, LPCWSTR lpszText1)
{
	ASSERT(IsValidType(nType));

	bool bRet = false;
	if (IsValidType(nType))
	{
		delete [] m_pDescriptions[nType];
		delete [] m_pDescriptions1[nType];
		m_pDescriptions[nType] = m_pDescriptions1[nType] = NULL;
		if (lpszText)
		{
			m_pDescriptions[nType] = new WCHAR[wcslen(lpszText) + 1];
			wcscpy(m_pDescriptions[nType], lpszText);
			bRet = true;
		}
		if (lpszText1)
		{
			m_pDescriptions1[nType] = new WCHAR[wcslen(lpszText1) + 1];
			wcscpy(m_pDescriptions1[nType], lpszText1);
			bRet = true;
		}
	}
	return bRet;
}

bool CDropDescription::SetInsert(LPCWSTR lpszInsert)
{
	delete [] m_lpszInsert;
	m_lpszInsert = NULL;
	if (lpszInsert)
	{
		m_lpszInsert = new WCHAR[wcslen(lpszInsert) + 1];
		wcscpy(m_lpszInsert, lpszInsert);
	}
	return NULL != m_lpszInsert;
}

// Get text string for specified image type.
// When b1 is true, return the string that includes the szInsert string when present.
LPCWSTR CDropDescription::GetText(DROPIMAGETYPE nType, bool b1) const
{
	return IsValidType(nType) ? 
		((b1 && m_pDescriptions1[nType]) ? m_pDescriptions1[nType] : m_pDescriptions[nType]) : NULL;
}

// Set drop description data insert and message text.
// Returns true when szInsert or szMessage has been changed.
bool CDropDescription::SetDescription(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszMsg) const
{
	return CopyInsert(pDropDescription, m_lpszInsert) | CopyMessage(pDropDescription, lpszMsg);
}

// Set drop description data insert and message text according to image type.
// Returns true when szInsert or szMessage has been changed.
bool CDropDescription::SetDescription(DROPDESCRIPTION *pDropDescription, DROPIMAGETYPE nType) const
{
	return SetDescription(pDropDescription, GetText(nType, HasInsert()));
}

// Set drop description data message text according to image type.
// When szInsert of the description is not empty, copy the string that includes szInsert.
// Returns true when szMessage has been changed.
bool CDropDescription::CopyText(DROPDESCRIPTION *pDropDescription, DROPIMAGETYPE nType) const
{
	return CopyMessage(pDropDescription, GetText(nType, HasInsert(pDropDescription)));
}

// Copy string to drop description szMessage member.
// Returns true if strings has been changed.
bool CDropDescription::CopyMessage(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszMsg) const
{
	ASSERT(pDropDescription);
	return CopyDescription(pDropDescription->szMessage, ARRAYSIZE(pDropDescription->szMessage), lpszMsg);
}

// Copy string to drop description szInsert member.
// Returns true if strings has been changed.
bool CDropDescription::CopyInsert(DROPDESCRIPTION *pDropDescription, LPCWSTR lpszInsert) const
{
	ASSERT(pDropDescription);
	return CopyDescription(pDropDescription->szInsert, ARRAYSIZE(pDropDescription->szInsert), lpszInsert);
}

// Helper function for CopyMessage() and CopyInsert().
bool CDropDescription::CopyDescription(LPWSTR lpszDest, size_t nDstSize, LPCWSTR lpszSrc) const
{
	ASSERT(lpszDest);
	ASSERT(nDstSize > 0);

	bool bChanged = false;
	if (lpszSrc && *lpszSrc)
	{
		if (wcscmp(lpszDest, lpszSrc))
		{
			bChanged = true;
			wcsncpy_s(lpszDest, nDstSize, lpszSrc, _TRUNCATE);
		}
	}
	else if (lpszDest[0])
	{
		bChanged = true;
		lpszDest[0] = L'\0';
	}
	return bChanged;
}




// Helper class to create and get clipboard HTML format data.
// See HTML Clipboard Format, http://msdn.microsoft.com/en-us/library/Aa767917.aspx

CHtmlFormat::CHtmlFormat()
{
	m_nStartHtml = m_nEndHtml = -1;
	m_nSize = 0;
	m_lpszStartHtml = NULL;
}

// Create HTML format data in global memory.
// NOTES:
// - Input string must be valid HTML (using entities for reserved HTML chars '&"<>').
// - The input may optionally contain a HTML header and body tags.
//   If not present, they are added here (if pWnd is not NULL it is used to get font information).
//   If present they must met these requirements:
//   - If a html section is provided, head and body sections must be also present.
//   - The header should contain a charset definition for UTF-8:
//     <meta http-equiv="content-type" content="text/html; charset=utf-8">
HGLOBAL CHtmlFormat::SetHtml(LPCTSTR lpszHtml, CWnd *pWnd /*= NULL*/)
{
	ASSERT(lpszHtml && *lpszHtml);

	m_nSize = 0;

	LPCWSTR lpszWide = NULL;
	CString str;
#ifndef _UNICODE
	CStringW strWide;
#endif

	try
	{
		// Check for missing HTML sections.
		bool bHasHtml = FindNoCase(lpszHtml, _T("<html")) >= 0;
		bool bHasHead = FindNoCase(lpszHtml, _T("<head")) >= 0;
		bool bHasBody = FindNoCase(lpszHtml, _T("<body")) >= 0;
		
		// Preallocate string to be big enough for the content:
		// - HTML section if not provided (110 chars)
		// - HEAD section if not provided (87 chars without title tag)
		// - BODY tags if not provided (18 chars), optional with font style if CWnd passed
		// - Fragment markers (42 = 22 + 20 chars)
		// - Text
		str.Preallocate(512 + static_cast<int>(_tcslen(lpszHtml)));
		if (!bHasHtml)
		{
			// Constant string is splitted for better readibility.
//			str = _T("<!DOCTYPE html PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">\r\n");
			str = _T("<!DOCTYPE html PUBLIC \"-//W3C//DTD HTML 4.01//EN\" \"http://www.w3.org/TR/html4/strict.dtd\">\r\n");
			str += _T("<html>\r\n");
		}
		else if (!bHasHead || !bHasBody)
		{
			ASSERT(0);								// must have <head> and <body> sections with <html> tag
			return NULL;
		}
		if (!bHasHead)
		{
			// Constant string is splitted for better readibility.
			str += _T("<head>\r\n");
			str += _T("<meta http-equiv=\"content-type\" content=\"text/html; charset=utf-8\">\r\n"); 
//			str += _T("<title>MyApp generated data</title>\r\n");
			str += _T("</head>\r\n");
		}
		else if (FindNoCase(lpszHtml, _T("</head>")) < 0)
		{
			ASSERT(0);								// missing closing head tag
			return NULL;
		}
		if (!bHasBody)
		{
			if (pWnd)
			{
				CString strFontStyle;
				GetFontStyle(strFontStyle, pWnd);
				str.AppendFormat(_T("<body style=\"%s\">\r\n"), strFontStyle.GetString());
			}
			else
				str += _T("<body>\r\n");
			str += _T("<!--StartFragment-->\r\n");
		}
		str += lpszHtml;
		if (!bHasBody)
		{
			str += _T("<!--EndFragment-->\r\n");
			str += _T("</body>\r\n");
		}
		else
		{
			// Insert fragment markers to mark the complete content between the body tags.
			int nPos = FindNoCase(str, _T("<body"));	// <body> tag
			nPos = str.Find(_T('>'), nPos);			// end of <body> tag
			do										// skip trailing white spaces
				++nPos;
			while (nPos < str.GetLength() && str[nPos] <= _T(' '));
			str.Insert(nPos, _T("<!--StartFragment-->\r\n"));

			nPos = FindNoCase(str, _T("</body>"), nPos);	// before </body> tag
			if (nPos < 0)
			{
				ASSERT(0);							// missing closing BODY tag
				return NULL;
			}
			str.Insert(nPos, _T("<!--EndFragment-->\r\n"));
		}
		if (!bHasHtml)
			str += _T("</html>\r\n");
#ifdef _UNICODE
		lpszWide = str.GetString();
#else
		// Convert input to wide char. Required for conversion to UTF-8.
		strWide = str.GetString();					// pass LPCSTR here to perform the conversion
		lpszWide = strWide.GetString();
#endif
	}
	catch (CMemoryException *pe)
	{
		pe->Delete();
	}

	HGLOBAL hGlobal = NULL;
	if (lpszWide)
	{
		// Get length of UTF-8 string including terminating NULL char.
		// The MSDN does not mention if the string should be NULL terminated or not.
		// Providing the NULL terminator is safe and always better.
		int nUtf8Size = ::WideCharToMultiByte(CP_UTF8, 0, lpszWide, -1, NULL, 0, NULL, NULL);
		if (nUtf8Size < 1)
		{
			TRACE0("CHtmlFormat::SetHtml: UTF-8 conversion failed\n");
			ASSERT(0);
			return false;
		}
		// Allocate buffer for HTML format descriptor and UTF-8 content. 
		// NOTE: Adjust length definition when changing the descriptor!
		const int nDescLen = 105;
		m_nSize = nDescLen + nUtf8Size;
		hGlobal = ::GlobalAlloc(GMEM_MOVEABLE, m_nSize);
		if (NULL != hGlobal)
		{
			// Convert to UTF-8
			bool bErr = false;
			LPSTR lpszBuf = static_cast<LPSTR>(::GlobalLock(hGlobal));
			LPSTR lpszUtf8 = lpszBuf + nDescLen;
			if (::WideCharToMultiByte(CP_UTF8, 0, lpszWide, -1, lpszUtf8, nUtf8Size, NULL, NULL) <= 0)
			{
				TRACE0("CHtmlFormat::SetHtml: UTF-8 conversion failed\n");
				bErr = true;
			}
			else
			{
				// Get the fragment marker positions
				LPCSTR lpszStartFrag = strstr(lpszUtf8, "<!--StartFragment-->");
				LPCSTR lpszEndFrag = strstr(lpszUtf8, "<!--EndFragment-->");
				ASSERT(lpszStartFrag);
				ASSERT(lpszEndFrag);
				// Adjust pointer to be behind inserted string followed by CR-LF
				lpszStartFrag += strlen("<!--StartFragment-->") + 2;

				// Create descriptor and prepend it to the string.
				// Use only ASCII chars here!
				// Using leading zeroes with the offsets here to avoid iterative calculations.
				// With valid nDescLen, _snprintf() will not append the NULL terminator here.
				VERIFY(nDescLen == _snprintf(
					lpszBuf, nDescLen,
					"Version:1.0\r\nStartHTML:%010d\r\nEndHTML:%010d\r\nStartFragment:%010d\r\nEndFragment:%010d\r\n",
					nDescLen, 
					nDescLen + nUtf8Size - 1,		// offset to next character behind string
					nDescLen + static_cast<int>(lpszStartFrag - lpszUtf8), 
					nDescLen + static_cast<int>(lpszEndFrag - lpszUtf8)));
			}
			::GlobalUnlock(hGlobal);
			if (bErr)
			{
				::GlobalFree(hGlobal);
				hGlobal = NULL;
				m_nSize = 0;
			}
		}
	}
	return hGlobal;
}

// Get HTML format content as allocated UTF-8 string.
// If bAll is true and the StartHTML and EndHTML keywords are valid, the complete HTML content is copied.
// Otherwise, the marked fragments are copied.
// NOTES:
//  - Supports multiple fragment markers.
//  - Uses fragment offsets from header (does not serch for markers)
//  - Does not use selection markers (uses fragment markers only).
// Searching for the markers would be a better solution when offsets are wrong!
LPSTR CHtmlFormat::GetUtf8(LPCSTR lpData, int nSize, bool bAll)
{
	ASSERT(lpData);
	ASSERT(nSize);

	// Clear member vars if called again.
	m_nSize = nSize;
	m_nStartHtml = m_nEndHtml = -1;
	m_arStartFrag.RemoveAll();
	m_arEndFrag.RemoveAll();
	m_lpszStartHtml = lpData + nSize - 1;
	m_strVersion.Empty();
	m_strSrcURL.Empty();

	// Get StartHTML and EndHTML offsets.
	// Note that these may be set to -1 in the descriptor.
	// Then try to locate the "<html>" tag.
	LPCSTR s = GetHtmlDescriptorValue(lpData, "StartHTML:", m_nStartHtml);
	if (NULL == s || m_nStartHtml < 0 || m_nStartHtml >= nSize)
	{
		TRACE0("CHtmlFormat::ParseHtmlUtf8: StartHTML not present or invalid\n");
		s = FindHtmlDescriptor(lpData, "<html");
		if (NULL != s)
			m_nStartHtml = static_cast<int>(s - lpData);
	}
	if (NULL != s)
	{
		m_lpszStartHtml = lpData + m_nStartHtml;
		s = GetHtmlDescriptorValue(lpData, "EndHTML:", m_nEndHtml);
		if (NULL == s || m_nEndHtml <= m_nStartHtml || m_nEndHtml > nSize)
		{
			TRACE0("CHtmlFormat::ParseHtmlUtf8: EndHTML not present or invalid\n");
			m_nEndHtml = nSize;
		}
		// EndHtml may be the index of the last included char or the next one.
		// This is not clearly stated in the MSDN.
		// If the indexed char is not NULL and the index is less than the buffer size,
		//  assume that it is the index of the last included char and add one.
		else if (m_nEndHtml < nSize && lpData[m_nEndHtml])
			++m_nEndHtml;
	}
	else if (bAll)
	{
		// StartHTML and EndHTML are not present or -1 and <html> tag not found.
		TRACE0("CHtmlFormat::ParseHtmlUtf8: Failed to detect HTML section. Can't copy all.\n");
		bAll = false;
	}

	int nLen = bAll ? m_nEndHtml - m_nStartHtml : 0;
	if (!bAll)
	{
		s = lpData;
		// There may be multiple fragment markers.
		do
		{
			int nStart, nEnd;
			s = GetHtmlDescriptorValue(s, "StartFragment:", nStart);
			if (NULL == s)
			{
#ifdef _DEBUG
				if (0 == m_arStartFrag.GetCount())
					TRACE0("CHtmlFormat::ParseHtmlUtf8: No Fragment markers found\n");
#endif
				break;
			}
			s = GetHtmlDescriptorValue(s, "EndFragment:", nEnd);
			if (NULL == s)
			{
				TRACE0("CHtmlFormat::ParseHtmlUtf8: Missing matching EndFragment marker\n");
				break;
			}
			if (nEnd >= nStart && nEnd <= nSize)
			{
				// nEnd may be the index of the last included char or the next one.
				// This is not clearly stated in the MSDN.
				// If the index points one char before the end fragment marker, increment it here
				//  to set it to the end marker. 
				if (0 == strncmp(lpData + nEnd + 1, "<!--EndFragment", 15))
					++nEnd;
				m_arStartFrag.Add(nStart);
				m_arEndFrag.Add(nEnd);
				nLen += nEnd - nStart;
				// No start of HTML detected so far.
				// Limit this while loop to the first fragment position.
				if (m_lpszStartHtml > lpData + nStart)
					m_lpszStartHtml = lpData + nStart;
			}
#ifdef _DEBUG
			else
			{
				TRACE0("CHtmlFormat::ParseHtmlUtf8: Invalid StartFragment / EndFragment marker\n");
				TRACE3(" Start: %d, End: %d, content length: %d\n", nStart, nEnd, nSize);
			}
#endif
		} while (s < m_lpszStartHtml);
	}

	LPSTR lpszHtml = NULL;
	try 
	{
		if (nLen)
		{
			lpszHtml = new char[nLen + 1];
			// Copy the complete HTML content including headers but without description.
			if (bAll)
				::CopyMemory(lpszHtml, lpData + m_nStartHtml, nLen);
			// Copy marked sections only.
			else
			{
				LPSTR lpszDst = lpszHtml;
				for (int i = 0; i < m_arStartFrag.GetCount(); i++)
				{
					int nPartLen = m_arEndFrag.GetAt(i) - m_arStartFrag.GetAt(i);
					::CopyMemory(lpszDst, lpData + m_arStartFrag.GetAt(i), nPartLen);
					lpszDst += nPartLen;
				}
				ASSERT(lpszDst == lpszHtml + nLen);
			}
			lpszHtml[nLen] = '\0';
		}
		// Get HTML format version string.
		GetHtmlDescriptorValue(lpData, "Version:", m_strVersion);
		// Get optional source URL string.
		GetHtmlDescriptorValue(lpData, "SourceURL:", m_strSrcURL);
	}
	catch (CMemoryException *pe)
	{
		pe->Delete();
		delete [] lpszHtml;
		lpszHtml = NULL;
	}
	return lpszHtml;
}

// Get HTML format content as allocated string.
LPTSTR CHtmlFormat::GetHtml(LPCSTR lpData, int nSize, bool bAll)
{
	LPTSTR lpszHtml = NULL;
	LPSTR lpszUtf8 = GetUtf8(lpData, nSize, bAll);
	if (lpszUtf8)
	{
#ifdef _UNICODE
		lpszHtml = CDragDropHelper::MultiByteToWideChar(lpszUtf8, -1, CP_UTF8);
#else
		lpszHtml = CDragDropHelper::MultiByteToMultiByte(lpszUtf8, -1, CP_UTF8);
#endif
		delete [] lpszUtf8;
	}
	return lpszHtml;
}

int CHtmlFormat::FindNoCase(LPCTSTR lpszBuffer, LPCTSTR lpszKeyword) const
{
	size_t nLen = _tcslen(lpszKeyword);
	LPCTSTR lpszFound = lpszBuffer;
	while (*lpszFound)
	{
		if (0 == _tcsnicmp(lpszFound, lpszKeyword, nLen))
			return static_cast<int>(lpszFound - lpszBuffer);
		++lpszFound;
	}
	return -1;
}

int CHtmlFormat::FindNoCase(const CString& str, LPCTSTR lpszKeyword, int nStart /*= 0*/) const
{
	int nStrLen = str.GetLength();
	size_t nKeyLen = _tcslen(lpszKeyword);
	while (nStart < nStrLen)
	{
		if (0 == _tcsnicmp(str.GetString() + nStart, lpszKeyword, nKeyLen))
			return nStart;
		++nStart;
	}
	return -1;
}

// Find keyword in buffer up to m_lpszStartHtml.
// Returns pointer to value if found.
LPCSTR CHtmlFormat::FindHtmlDescriptor(LPCSTR lpszBuffer, LPCSTR lpszKeyword) const
{
	size_t nLen = strlen(lpszKeyword);
	LPCSTR lpszFound = lpszBuffer;
	while (lpszFound < m_lpszStartHtml)
	{
		if (0 == _strnicmp(lpszFound, lpszKeyword, nLen))
			return lpszFound + nLen;
		++lpszFound;
	}
	return NULL;
}

// Get integer keyword value.
LPCSTR CHtmlFormat::GetHtmlDescriptorValue(LPCSTR lpszBuffer, LPCSTR lpszKeyword, int& nValue) const
{
	LPCSTR lpszFound = FindHtmlDescriptor(lpszBuffer, lpszKeyword);
	if (lpszFound)
		nValue = atoi(lpszFound);
	return lpszFound;
}

// Get string keyword value.
LPCSTR CHtmlFormat::GetHtmlDescriptorValue(LPCSTR lpszBuffer, LPCSTR lpszKeyword, CStringA& strValue) const
{
	LPCSTR lpszFound = FindHtmlDescriptor(lpszBuffer, lpszKeyword);
	if (lpszFound)
	{
		LPCSTR lpszEnd = strpbrk(lpszFound, "\r\n");
		if (lpszEnd)
			strValue.SetString(lpszFound, static_cast<int>(lpszEnd - lpszFound));
	}
	return lpszFound;
}

// Get mapped font name for MS Shell Dlg logical fonts.
// 'MS Shell Dlg' and 'MS Shell Dlg 2' fonts are logical fonts that map to other fonts
//  according to localization. 
// See http://msdn.microsoft.com/en-us/library/windows/desktop/dd374112%28v=vs.85%29.aspx
/*static*/ LPCTSTR CHtmlFormat::GetMappedFontName(const LOGFONT& lf)
{
	LPCTSTR lpszName = NULL;
	if (0 == _tcscmp(lf.lfFaceName, _T("MS Shell Dlg 2")))
	{
		// 'MS Shell Dlg 2' is used with Windows 2000 and later and is always mapped to 'Tahoma'.
		lpszName = _T("Tahoma");
	}
	else if (0 == _tcscmp(lf.lfFaceName, _T("MS Shell Dlg")))
	{
		// 'MS Shell Dlg' is mapped to 'Microsoft Sans Serif' if the language 
		//  is not Japanese. With Japanese, 'MS UI Gothic' is used when:
		// Vista or later: Machine default language is Japanese regardless of the install language.
		// XP MUI: System locale and UI language are set to Japanese.
		// 2000, XP: Install language is Japanese. 
		// See the above link for older Windows versions.
		if (LANG_JAPANESE == PRIMARYLANGID(LANGIDFROMLCID(::GetSystemDefaultLCID())))
			lpszName = _T("MS UI Gothic");
		else
			lpszName = _T("Microsoft Sans Serif");
	}
	return lpszName;
}

// Get HTML font style string from CWnd font.
/*static*/ void CHtmlFormat::GetFontStyle(CString& strStyle, CWnd *pWnd)
{
	ASSERT_VALID(pWnd);

	LOGFONT lf;
	CFont *pFont = pWnd->GetFont();
	pFont->GetLogFont(&lf);
	LPCTSTR lpszFontName = GetMappedFontName(lf);
	if (NULL != lpszFontName)
		strStyle.Format(_T("font-family:'%s',sans-serif"), lpszFontName);
	else
	{
		strStyle.Format(_T("font-family:'%s',"), lf.lfFaceName);
		switch (lf.lfPitchAndFamily & 0xF0)
		{
		case FF_ROMAN : strStyle += _T("serif"); break;
//		case FF_SWISS : strStyle += _T("sans-serif"); break;
		case FF_MODERN : strStyle += _T("monospace"); break;
		case FF_SCRIPT : strStyle += _T("cursive"); break;
		case FF_DECORATIVE : strStyle += _T("fantasy"); break;
		default : strStyle += _T("sans-serif");
		}
	}
	if (lf.lfItalic)
		strStyle += _T("; font-style:italic");
	if (lf.lfWeight > 0 && lf.lfWeight != 400)		// weight 400 is normal
		strStyle.AppendFormat(_T("; font-weight:%d"), lf.lfWeight);
	if (lf.lfHeight < 0)
	{
		strStyle.AppendFormat(_T("; font-size:%.1fpt"), 
			-72.0 * lf.lfHeight / ::GetDeviceCaps(pWnd->GetDC()->GetSafeHdc(), LOGPIXELSY));
	}
	if (lf.lfUnderline)
		strStyle += _T("; text-decoration:underline");
	if (lf.lfStrikeOut)
		strStyle += _T("; text-decoration:line-through");
}


CTextToRTF::CTextToRTF()
{
	m_bDelim = false;
	m_lpszFormat[0] = _T('\0');
}

// Convert single wide char to RTF.
// Returns pointer to expansion string and prints format string.
//
// Return value is a control word string when cWide is to be replaced by a control word.
// If cWide is a reserved character, the string contains the backslash character. 
// If neither applies, NULL is returned.
//
// m_lpszFormat is set to the input char cWide if it is an ASCII character (including reserved chars).
// If cWide is an ANSI or Unicode character, lpszFormat is set to the corresponding encoding control sequence.
// Otherwise, lpszFormat is set to an empty string.
LPCTSTR CTextToRTF::CharToRtf(wchar_t cWide)
{
	LPCTSTR lpszControl = NULL;						// expansion string to be returned
	m_lpszFormat[0] = _T('\0');						// clear formatted string
	switch (cWide)
	{
	case L'\t' : lpszControl = _T("\\tab"); break;	// tab
	case L'\n' : lpszControl = _T("\\par"); break;	// end of paragraph
	case L'\f' : lpszControl = _T("\\page"); break;	// page break
	case 0x00A0 : lpszControl = _T("\\~"); break;	// nonbreaking space
	case 0x2002 : lpszControl = _T("\\enspace"); break;
	case 0x2003 : lpszControl = _T("\\emspace"); break;
	case 0x2011 : lpszControl = _T("\\_"); break;	// nonbreaking hyphen
	case 0x2013 : lpszControl = _T("\\endash"); break;
	case 0x2014 : lpszControl = _T("\\emdash"); break;
	case 0x2018 : lpszControl = _T("\\lquote"); break;
	case 0x2019 : lpszControl = _T("\\rquote"); break;
	case 0x201C : lpszControl = _T("\\ldblquote"); break;
	case 0x201D : lpszControl = _T("\\rdblquote"); break;
	case 0x2022 : lpszControl = _T("\\bullet"); break;
	case L'\\' : 
	case L'{' : 
	case L'}' : 
		lpszControl = _T("\\");						// escape reserved chars; fall through to default
	default :
		if (cWide >= 0x80)							// not an ASCII character
		{
			char lpszMB[MB_LEN_MAX];				// output buffer for WideCharToMultiByte()
			// Check if it is an ANSI character from the current code page
            if (1 == ::WideCharToMultiByte(CP_THREAD_ACP, 0, &cWide, 1, lpszMB, sizeof(lpszMB), NULL, NULL)) {
                // x-studio365 spec: pitfalls: fix bug, replace 'sizeof(m_lpszFormat)' by ARRAYSIZE(m_lpszFormat)
                _stprintf_s(m_lpszFormat, ARRAYSIZE(m_lpszFormat)/*sizeof(m_lpszFormat)*/, _T("\\'%x"), static_cast<unsigned char>(lpszMB[0]));
            }
			else									// it is an Unicode character
			{
				// Try to get a composite character.
				if (1 == ::WideCharToMultiByte(CP_THREAD_ACP, WC_COMPOSITECHECK, &cWide, 1, lpszMB, sizeof(lpszMB), NULL, NULL))
				{
					// The composite character must be ASCII and non-digit.
					if ((lpszMB[0] & 0x80) || isdigit(lpszMB[0]))
						lpszMB[0] = '?';
				}
				else
					lpszMB[0] = '?';
				// RTF numeric parameters are signed 16-bit values.
				// So use the format "%hd" to print values >= 0x8000 as signed numbers.
				// We must use the 'h' prefix here because the wchar_t type is unsigned short.
                // x-studio365 spec: pitfalls: fix bug, replace 'sizeof(m_lpszFormat)' by ARRAYSIZE(m_lpszFormat)
				_stprintf_s(m_lpszFormat, ARRAYSIZE(m_lpszFormat)/*sizeof(m_lpszFormat)*/, _T("\\u%hd%hc"), cWide, lpszMB[0] ? lpszMB[0] : '?');
			}
		}
		else if (cWide >= L' ')						// ignore all other control characters
		{											//  they would be ignored by RTF readers
			// If the previous char has been expanded to a control word, it must
			//  be delimited by a space if this char is a letter, a digit, or a space.
			if (m_bDelim && (L' ' == cWide || iswalnum(cWide)))
				lpszControl = _T(" ");
#ifdef _UNICODE
			m_lpszFormat[0] = cWide;
#else
			m_lpszFormat[0] = static_cast<char>(cWide);	// cast to char (cWide contains an ASCII char here)
#endif
			m_lpszFormat[1] = _T('\0');
		}
	}
	return lpszControl;
}

#ifndef _UNICODE
// Convert multi-byte char to wide char and return pointer to next char.
// Helper function for text to RTF conversion with non-Unicode builds.
int CTextToRTF::ToWideChar(LPCSTR lpszText, wchar_t& cWide)
{
	cWide = L'\0';
	int n = 0;
	// Check for lead byte with double-byte character sets.
	if (::IsDBCSLeadByteEx(CP_THREAD_ACP, (BYTE)*lpszText))
	{
		if (1 == ::MultiByteToWideChar(CP_THREAD_ACP, 0, lpszText, 2, &cWide, 1))
			n = 2;
	}
	if (n <= 0)
	{
		// Assume single-byte character set (ANSI).
		// This will fail with multi-byte character sets (e.g. UTF-8).
		n = ::MultiByteToWideChar(CP_THREAD_ACP, 0, lpszText, 1, &cWide, 1);
		// Fall back to CRT conversion.
		// NOTE: This uses the locale set by setlocale() which may be different from those
		//  used by the thread!
		if (n <= 0)
			n = mbtowc(&cWide, lpszText, MB_CUR_MAX);
	}
	ASSERT(n > 0);								// *lpszText is null char or invalid multi byte char
	return n;
}
#endif

// Convert text string to RTF.
// strRtf:   CString receiving the converted string.
// lpszText: Text to be converted.
// NOTE: 
//  With Unicode builds, the output string must be converted to char to be valid RTF!
void CTextToRTF::StringToRtf(CString& strRtf, LPCTSTR lpszText)
{
	ASSERT(lpszText && *lpszText);

	m_bDelim = false;								// control word delimiter flag

#if 0
	size_t nSize = 2 * _tcslen(lpszText);			// guess the size
#else
	// Calculate the resulting string size.
	size_t nSize = 0;
	LPCTSTR s = lpszText;
	while (*s)
	{
        if (nSize == (41 - 9))
        {
            int brk = 0;
        }
#ifdef _UNICODE
		LPCTSTR lpszControl = CharToRtf(*s++);
#else
		wchar_t cWide = L'\0';
		int n = ToWideChar(s, cWide);
		if (n <= 0)
			break;
		s += n;
		LPCTSTR lpszControl = CharToRtf(cWide);
#endif
		if (lpszControl)
		{
			nSize += _tcslen(lpszControl);
			// If this char is expanded to a control word, it must be separated from the next char
			//  when that is a letter, a digit, or a space.
			// Note that this does not apply to control symbols (backslash followed by a single char).
			m_bDelim = _T('\\') == lpszControl[0] && lpszControl[1] && lpszControl[2];
		}
		else
			m_bDelim = false;

		nSize += _tcslen(m_lpszFormat);
        pChecker->Preallocate(16);
	}
	m_bDelim = false;
#endif

	try
	{
        strRtf = _T("");
		strRtf.Preallocate(static_cast<int>(nSize));
		
		while (*lpszText)
		{
#ifdef _UNICODE
			LPCTSTR lpszControl = CharToRtf(*lpszText++);
#else
			wchar_t cWide = L'\0';
			int n = ToWideChar(lpszText, cWide);
			if (n <= 0)
				break;
			lpszText += n;
			LPCTSTR lpszControl = CharToRtf(cWide);
#endif
			if (lpszControl)
			{
				strRtf += lpszControl;
				// If this char is expanded to a control word, it must be separated from the next char
				//  when that is a letter, a digit, or a space.
				// Note that this does not apply to control symbols (backslash followed by a single char).
				m_bDelim = _T('\\') == lpszControl[0] && lpszControl[1] && lpszControl[2];
			}
			else
				m_bDelim = false;
			if (m_lpszFormat[0])
				strRtf += m_lpszFormat;
		}
	}
	catch (CMemoryException *pe)
	{
		pe->Delete();
	}
}

// Convert plain text to RTF including RTF header.
// strRtf:       CString receiving the converted string.
// lpszText:     Text to be converted.
// lpszFontName: Optional name of font.
// nFontSize:    Optional font size in half-points (default is 24 when not specified).
// Returns pointer to the RTF string upon success and NULL if out of memory.
// NOTE: 
//  With Unicode builds, the output string must be converted to char to be valid RTF!
LPCTSTR CTextToRTF::TextToRtf(CString& strRtf, LPCTSTR lpszText, LPCTSTR lpszFontName /*= NULL*/, int nFontSize /*= 0*/)
{
	ASSERT(lpszText && *lpszText);

    pChecker = &strRtf;

	CString strContent;
	StringToRtf(strContent, lpszText);				// create RTF content string
	if (strContent.IsEmpty())
		return NULL;

	try
	{
		strRtf.Preallocate(128 + strContent.GetLength());
		strRtf = _T("{\\rtf1\\ansi");				// RTF header
		LCID nLCID = ::GetThreadLocale();			// get default ANSI code page for thread locale
		UINT nCP = CDragDropHelper::GetCodePage(nLCID, false);
		if (nCP >= 437 && nCP <= 1361)				// all ANSI/OEM code pages supported by RTF are in this range
			strRtf.AppendFormat(_T("\\ansicpg%d"), nCP);
		if (lpszFontName && *lpszFontName)
		{
			strRtf.AppendFormat(
				_T("\\deff0\\deflang%u{\\fonttbl{\\f0\\fnil\\fcharset0 %s;}}"),
				nLCID, lpszFontName);
		}
		strRtf += _T("\\uc1\\pard");				// Unicode encoding, reset paragraph settings
		if (lpszFontName && *lpszFontName)
		{
			strRtf += _T("\\f0");					// font selection
			if (nFontSize > 0)						// font size in half-points (default is 24)
				strRtf.AppendFormat(_T("\\fs%d"), nFontSize);
		}
//		if (_T(' ') == strContent.GetAt(0) || _istalnum(strContent.GetAt(0)))
			strRtf += _T(' ');						// control word delimiter
		strRtf += strContent;						// append the text
		strRtf += _T('}');							// block terminator
	}
	catch (CMemoryException *pe)
	{
		pe->Delete();
		strRtf.Empty();
	}
	return strRtf.IsEmpty() ? NULL : strRtf.GetString();
}

// Text to RTF with font name and size from CWnd
LPCTSTR CTextToRTF::TextToRtf(CString& strRtf, LPCTSTR lpszText, CWnd *pWnd)
{
	if (pWnd)
	{
		ASSERT_VALID(pWnd);
		return TextToRtf(strRtf, lpszText, pWnd->GetFont(), pWnd->GetDC()->GetSafeHdc());
	}
	return TextToRtf(strRtf, lpszText, L"C", 1);
}

// Text to RTF with font name and size from CFont
LPCTSTR CTextToRTF::TextToRtf(CString& strRtf, LPCTSTR lpszText, CFont *pFont, HDC hDC /*= NULL*/)
{
	ASSERT_VALID(pFont);

	int nFontSize = 0;
	LPCTSTR lpszFontName = NULL;
	LOGFONT lf;
	if (pFont && pFont->GetLogFont(&lf))
	{
		lpszFontName = CHtmlFormat::GetMappedFontName(lf);
		if (NULL == lpszFontName)
			lpszFontName = lf.lfFaceName;
		// Font size. Note that RTF expects the size in half-points.
		// So the size must be doubled here.
		if (lf.lfHeight < 0)
		{
			bool bHasDC = (hDC != NULL);
			if (!bHasDC)
				hDC = ::GetDC(NULL);
			nFontSize = static_cast<int>(0.5 - ((144.0 * lf.lfHeight) / ::GetDeviceCaps(hDC, LOGPIXELSY)));
			if (!bHasDC)
				::ReleaseDC(NULL, hDC);
		}
	}
	return TextToRtf(strRtf, lpszText, lpszFontName, nFontSize);
}
