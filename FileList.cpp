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

// FileList.cpp : 
// A ClistCtrl (CMyListCtrl) derived class to list the files in directories
//  in a similar way like the Windows Explorer.
// I was just too lazy to create my own icons for an example list control.
// With such file lists, the system provides small and large icons.

#include "stdafx.h"
#include "OleDataDemo.h"
#include "FileList.h"
#include ".\filelist.h"

#include <lmcons.h>									// UNLEN, DNLEN definitions
#include <aclapi.h>									// GetNamedSecurityInfo() function
#include <Sddl.h.>									// ConvertSidToStringSid() function

#pragma comment(lib, "advapi32.lib")				// Access control functions

#define SORT_BUF_SIZE	256

// CFileList

IMPLEMENT_DYNAMIC(CFileList, CMyListCtrl)
CFileList::CFileList()
{
	m_bSortDesc = false;							// sort direction
	m_nSortCol = FileColumns::NameColumn;			// sort column
	m_bDragHighLight = false;						// CMyListCtrl base class member
#ifdef CACHED_GET_OWNER_DATA
	m_lpszComputerName[0] = _T('\0');
	::ZeroMemory(m_pSidBuf, sizeof(m_pSidBuf));
#endif
}

CFileList::~CFileList()
{
}


BEGIN_MESSAGE_MAP(CFileList, CMyListCtrl)
	ON_NOTIFY_REFLECT(LVN_COLUMNCLICK, OnLvnColumnclick)
	ON_NOTIFY_REFLECT(NM_DBLCLK, OnNMDblclk)
END_MESSAGE_MAP()


void CFileList::SetColumns()
{
	if (0 == GetHeaderCtrl()->GetItemCount())
	{
		InsertColumn(FileColumns::NameColumn, _T("Name"), LVCFMT_LEFT, 120, FileColumns::NameColumn);
		InsertColumn(FileColumns::ModifiedColumn, _T("Modified"), LVCFMT_LEFT, 120, FileColumns::ModifiedColumn);
//		InsertColumn(FileColumns::CreatedColumn, _T("Created"), LVCFMT_LEFT, 0, FileColumns::CreatedColumn);
//		InsertColumn(FileColumns::AccessedColumn, _T("Accessed"), LVCFMT_LEFT, 0, FileColumns::AccessedColumn);
		InsertColumn(FileColumns::TypeColumn, _T("Type"), LVCFMT_LEFT, 100, FileColumns::TypeColumn);
		InsertColumn(FileColumns::SizeColumn, _T("Size"), LVCFMT_RIGHT, 60, FileColumns::SizeColumn);
		InsertColumn(FileColumns::OwnerColumn, _T("Owner"), LVCFMT_LEFT, 100, FileColumns::OwnerColumn);
		InsertColumn(FileColumns::AttributeColumn, _T("Attributes"), LVCFMT_LEFT, 60, FileColumns::AttributeColumn);
	}
}

// Populate list with folder content specified by name.
// Returns number of items added to the list.
int CFileList::SetPath(LPCTSTR lpszPath)
{
	ASSERT(lpszPath);

	size_t nLen = _tcslen(lpszPath);
	if (nLen && _T('\\') == lpszPath[nLen-1])
		--nLen;
	m_strPath.SetString(lpszPath, static_cast<int>(nLen));
	return ParseDirectory();
}

// Populate list with folder content specified by CSIDL.
// Returns number of items added to the list.
int CFileList::SetPath(int nFolder)
{
	int nItems = 0;
	TCHAR lpszPath[MAX_PATH];
	if (SUCCEEDED(::SHGetFolderPath(NULL, nFolder, NULL, SHGFP_TYPE_CURRENT, lpszPath)))
		nItems = SetPath(lpszPath);
	return nItems;
}

// Populate list with folder content specified by m_strPath.
// Returns number of items added to the list.
int CFileList::ParseDirectory()
{
	ASSERT(!m_strPath.IsEmpty());

	if (m_strPath.IsEmpty())
		return 0;

	CWaitCursor wait;

	SetColumns();									// Add columns with first call
	SetRedraw(0);									// Disable redrawing while updating the list
	DeleteAllItems();								// Remove all items from previous call
	m_LargeImages.DeleteImageList();				// Delete image lists from previous call
	m_SmallImages.DeleteImageList();

	int nItems = 0;
	CString strItem;
	CString strFind(m_strPath);
	strFind += _T("\\*.*");
	CFileFindEx finder;
	BOOL bFound = finder.FindFile(strFind.GetString());
	while (bFound)
	{
		bFound = finder.FindNextFile();

		// Skip this directory (".")
		if (_T(".") == finder.GetFileName())
			continue;

		// Skip links
		SHFILEINFO fi;
		if (::SHGetFileInfo(finder.GetFilePath().GetString(), 0, &fi, sizeof(fi), SHGFI_ATTRIBUTES))
		{
			if (fi.dwAttributes & SFGAO_LINK)
				continue;
		}

		// Insert item and set file/directory name.
		int nItem = InsertItem(nItems, finder.GetFileName().GetString());

		// Set item data to file attributes.
		// These are used below when getting the file type icons and are then
		//  replaced by the item index with directory flag.
		VERIFY(SetItemData(nItem, finder.GetAttributes()));

		// Date/time stamps
		CTime time;
		if (finder.GetLastWriteTime(time))
		{
			SetItemText(nItem, FileColumns::ModifiedColumn, 
				time.Format(_T("%Y-%m-%d %H:%M:%S")).GetString());
		}
#if 0
		if (finder.GetLastAccessTime(time))
		{
			SetItemText(nItem, FileColumns::AccessedColumn, 
				time.Format(_T("%Y-%m-%d %H:%M:%S")).GetString());
		}
		if (finder.GetCreationTime(time))
		{
			SetItemText(nItem, FileColumns::CreatedColumn, 
				time.Format(_T("%Y-%m-%d %H:%M:%S")).GetString());
		}
#endif

		// Attributes
		VERIFY(SetItemText(nItem, FileColumns::AttributeColumn, GetAttributeString(strItem, finder.GetAttributes())));

		// Owner
		if (GetOwner(strItem, finder.GetFilePath().GetString()))
			VERIFY(SetItemText(nItem, FileColumns::OwnerColumn, strItem.GetString()));

		// File size rounded up to full KB.
		if (!finder.IsDirectory())
		{
			strItem.Format(_T("%I64u KB"), (finder.GetLength() + 1023ULL) / 1024ULL);
			VERIFY(SetItemText(nItem, FileColumns::SizeColumn, strItem.GetString()));
		}
		++nItems;
	}
	finder.Close();

	// Get file type icons and file type string.
	// This is done inside another loop because we know now the number of items
	//  so that the image lists can be created using the final size avoiding reallocations.
	LVITEM lvi;
	lvi.iSubItem = 0;
	for (lvi.iItem = 0; lvi.iItem < nItems; lvi.iItem++)
	{
		// Get file tpye icons and file type string.
		GetPath(strItem, lvi.iItem);
		SHFILEINFO fiLarge = { };
		SHFILEINFO fiSmall = { };;
		DWORD dwAttr = static_cast<DWORD>(GetItemData(lvi.iItem));
		// By using SHGFI_USEFILEATTRIBUTES with passed file attribute, the file is not accessed.
		::SHGetFileInfo(strItem.GetString(), dwAttr, &fiLarge, sizeof(fiLarge), SHGFI_USEFILEATTRIBUTES | SHGFI_ICON | SHGFI_TYPENAME);
		::SHGetFileInfo(strItem.GetString(), dwAttr, &fiSmall, sizeof(fiSmall), SHGFI_USEFILEATTRIBUTES | SHGFI_ICON | SHGFI_SMALLICON);
		// File type string
		SetItemText(lvi.iItem, FileColumns::TypeColumn, fiLarge.szTypeName);
		// Large icon
		lvi.iImage = AddImage(m_LargeImages, fiLarge.hIcon, nItems);
		// Small icon
		AddImage(m_SmallImages, fiSmall.hIcon, nItems);
		// Set image index and data.
		// Data is set to the item index with directory flag.
		// The directory flag is used for fast detection of directories with sorting and by handlers.
		// The item index is used to sort by the original order.
		lvi.mask = (lvi.iImage >= 0) ? LVIF_PARAM | LVIF_IMAGE : LVIF_PARAM;
		lvi.lParam = (dwAttr & FILE_ATTRIBUTE_DIRECTORY) ? lvi.iItem | nDirectoryFlag : lvi.iItem;
		VERIFY(SetItem(&lvi));
	}
	if (nItems)
	{
		SetImageList(&m_LargeImages, LVSIL_NORMAL);
		SetImageList(&m_SmallImages, LVSIL_SMALL);
		Sort(m_nSortCol, false);
	}
	SetRedraw(1);
	return nItems;
}

// Create file attribute string.
LPCTSTR CFileList::GetAttributeString(CString& str, DWORD dwAttr) const
{
	const TCHAR pAttrChars[] = _T("RHS DA  T  CO E");
	int i = 0;
	str = _T("");
	while (dwAttr && pAttrChars[i])
	{
		if ((dwAttr & 1) && pAttrChars[i] > _T(' '))
			str += pAttrChars[i];
		dwAttr >>= 1;
		++i;
	}
	return str.GetString();
}

// Get full path for specified item.
void CFileList::GetPath(CString& strPath, int nItem) const
{
	ASSERT(nItem >= 0 && nItem < GetItemCount());

	strPath = m_strPath;
	strPath += _T('\\');
	strPath += GetItemText(nItem, FileColumns::NameColumn);
}

// Helper function to add an icon to a image list.
int CFileList::AddImage(CImageList& ImageList, HICON hIcon, int nInitial)
{
	int nImage = -1;
	if (hIcon)
	{
		if (NULL == ImageList.m_hImageList)			// First icon to be added to list
		{
			ICONINFO ii;
			VERIFY(::GetIconInfo(hIcon, &ii));
			// Get the size of the icon from it's bitmap.
			// With a black & white icon, ii.hbmColor is NULL and
			//  the mask bitmap contains the image and the XOR mask.
			BITMAP bm;
			VERIFY(::GetObject(ii.hbmColor ? ii.hbmColor : ii.hbmMask, sizeof(bm), &bm));
			if (NULL == ii.hbmColor)
				bm.bmHeight /= 2;
			VERIFY(ImageList.Create(bm.bmWidth, bm.bmHeight, bm.bmBitsPixel | ILC_MASK, nInitial, 1));
		}
		nImage = ImageList.Add(hIcon);
		VERIFY(::DestroyIcon(hIcon));
	}
	return nImage;
}

// Check if pSid is one of the builtin accounts.
BOOL CFileList::IsBuiltinSid(PSID pSid)
{
#ifdef CACHED_GET_OWNER_DATA
	BOOL bRet = ::IsValidSid(reinterpret_cast<PSID>(m_pSidBuf));
	if (!bRet)										// first call
	{
		DWORD dwSidSize = sizeof(m_pSidBuf);
		bRet = ::CreateWellKnownSid(WinBuiltinAdministratorsSid, NULL, reinterpret_cast<PSID>(m_pSidBuf), &dwSidSize);
	}
	if (bRet)
		bRet = ::EqualPrefixSid(pSid, reinterpret_cast<PSID>(m_pSidBuf));
#else
	BYTE pSidBuf[SECURITY_MAX_SID_SIZE];
	DWORD dwSidSize = sizeof(pSidBuf);
	// Create any SID for a builtin account (first parameter must be one of the WinBuiltin... enums).
	BOOL bRet = ::CreateWellKnownSid(WinBuiltinAdministratorsSid, NULL, reinterpret_cast<PSID>(pSidBuf), &dwSidSize);
	if (bRet)
		bRet = ::EqualPrefixSid(pSid, reinterpret_cast<PSID>(pSidBuf));
#endif
	return bRet;
}

// Check if lpszDomain is the local computer name.
BOOL CFileList::IsLocalDomain(LPCTSTR lpszDomain)
{
	BOOL bRet = ((NULL != lpszDomain) && *lpszDomain);
#ifdef CACHED_GET_OWNER_DATA
	if (bRet && !m_lpszComputerName[0])
	{
		DWORD dwSize = ARRAYSIZE(m_lpszComputerName);
		bRet = ::GetComputerName(m_lpszComputerName, &dwSize);
	}
	if (bRet)
		bRet = (0 == _tcsicmp(lpszDomain, m_lpszComputerName));
#else
	if (bRet)
	{
		TCHAR lpszComputerName[MAX_COMPUTERNAME_LENGTH + 1];
		DWORD dwSize = ARRAYSIZE(lpszComputerName) ;
		if (::GetComputerName(lpszComputerName, &dwSize))
			bRet = (0 == _tcsicmp(lpszDomain, lpszComputerName));
	}
#endif
	return bRet;
}

// Get owner account name for file or directory.
bool CFileList::GetOwner(CString& strOwner, LPCTSTR lpszFile, unsigned nFlags /*= (UINT)-1*/)
{
	strOwner = _T("");
	// GetNamedSecurityInfo() requires a non const LPTSTR as object name.
	LPTSTR lpszObjectName = _tcsdup(lpszFile);
	DWORD dwRes = (NULL == lpszObjectName) ? ERROR_NOT_ENOUGH_MEMORY : ERROR_SUCCESS;
	PSID pSidOwner = NULL;
	PSECURITY_DESCRIPTOR pSD = NULL;
	if (ERROR_SUCCESS == dwRes)
	{
		dwRes = ::GetNamedSecurityInfo(
			lpszObjectName,							// object name
			SE_FILE_OBJECT,							// object is file or directory
			OWNER_SECURITY_INFORMATION,				// get information about owner into pSidOwner 
			&pSidOwner,								// returned owner info
			NULL,									// unused: returned primary group info
			NULL,									// unused: returned DACL
			NULL,									// unused: returned SACL
			&pSD);									// returned security descriptor
		free(lpszObjectName);
	}

	if (ERROR_SUCCESS == dwRes && ::IsValidSid(pSidOwner))
	{
		TCHAR lpszAccount[UNLEN + 1];				// max. user/group name length (256)
		TCHAR lpszDomain[DNLEN + 1];				// max. domain name length (15)
		DWORD dwAccountSize = ARRAYSIZE(lpszAccount) ;
		DWORD dwDomainSize = ARRAYSIZE(lpszDomain) ;
		SID_NAME_USE SNU = SidTypeUnknown;
		BOOL bRes = ::LookupAccountSid(
			NULL,									// name of local or remote computer
			pSidOwner,								// security identifier
			lpszAccount,							// account name buffer
			&dwAccountSize,							// size of account name buffer 
			lpszDomain,								// domain name buffer
			&dwDomainSize,							// size of domain name buffer
			&SNU);									// SID type
		if (bRes)
		{
			// Check if name should be prefixed with the domain.
			bool bDomain = (lpszDomain[0] && (nFlags & GET_OWNER_DOMAIN));
			if (bDomain)
			{
				// Check if specific domains should be excluded.
				// Well known group accounts (mainly SYSTEM user)
				if ((nFlags & GET_OWNER_NO_WELLKNOWNGROUP) && SidTypeWellKnownGroup == SNU)
					bDomain = false;
				// Local domain (computer name)
				else if ((nFlags & GET_OWNER_NO_LOCAL) && IsLocalDomain(lpszDomain))
					bDomain = false;
				// BUILTIN accounts
				else if ((nFlags & GET_OWNER_NO_BUILTIN) && IsBuiltinSid(pSidOwner))
					bDomain = false;
			}
			if (bDomain)
			{
				strOwner = lpszDomain;
				strOwner += _T('\\');
			}
			strOwner += lpszAccount;
		}
		// Optional return the SID as string when the name could not be retrieved.
		// Possible errors:
		//  ERROR_INVALID_PARAMETER: Usually an invalid SID.
		//  ERROR_NONE_MAPPED: Deleted account or a remote account that can't be resolved.
		//  ERROR_TRUSTED_RELATIONSHIP_FAILURE
		else if ((nFlags & GET_OWNER_SIDSTR) && ERROR_INVALID_PARAMETER != ::GetLastError())
		{
			LPTSTR lpszOwner = NULL;
			if (::ConvertSidToStringSid(pSidOwner, &lpszOwner))
				strOwner = lpszOwner;
			::LocalFree(lpszOwner);
		}
	}
	// Free the returned security descriptor after all.
	// LookupAccountSid() may fail when this is freed before.
	::LocalFree(pSD);
	return !strOwner.IsEmpty();
}

// Change directory when double clicking on a directory item
void CFileList::OnNMDblclk(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE lpnmitem = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	if (IsDirectory(lpnmitem->iItem))
	{
		CString strDir = GetItemText(lpnmitem->iItem, FileColumns::NameColumn);
		if (_T("..") == strDir)
		{
			int nPos = m_strPath.ReverseFind(_T('\\'));
			if (nPos > 1)
				m_strPath.Truncate(nPos);
		}
		else
		{
			m_strPath += _T('\\');
			m_strPath += strDir;
		}
		ParseDirectory();
	}
	*pResult = 0;
}

// Sort by clicked column.
void CFileList::OnLvnColumnclick(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	Sort(pNMLV->iSubItem, true);
	*pResult = 0;
}

// Set sort arrow on header and sort the list.
// When passing a negative column, arrows are removed from the header and 
//  the items are shown in the order when the list has been populated.
// When bAutoDirection is true and the current sort column is identical to the passed column,
//  the sort direction is reversed.
void CFileList::Sort(int nSortCol, bool bAutoDirection)
{
	ASSERT(nSortCol < GetHeaderCtrl()->GetItemCount());

	HDITEM hdi;
	hdi.mask = HDI_FORMAT;
	CHeaderCtrl *pHeaderCtrl = GetHeaderCtrl();
	if (m_nSortCol >= 0)							// remove arrow from old sort column
	{												//  or set it when same column
		pHeaderCtrl->GetItem(m_nSortCol, &hdi);
		hdi.fmt &= ~(HDF_SORTDOWN | HDF_SORTUP);
		if (m_nSortCol == nSortCol)
		{
			if (bAutoDirection)
				m_bSortDesc = !m_bSortDesc;			// change direction
			hdi.fmt |= m_bSortDesc ? HDF_SORTDOWN : HDF_SORTUP;
		}
		pHeaderCtrl->SetItem(m_nSortCol, &hdi);
	}
	if (nSortCol >= 0 && nSortCol != m_nSortCol)	// add arrow to new sort column
	{
		pHeaderCtrl->GetItem(nSortCol, &hdi);
		hdi.fmt |= HDF_SORTUP;
		m_bSortDesc = false;						// always sort ascending when a new column
		pHeaderCtrl->SetItem(nSortCol, &hdi);
	}
	m_nSortCol = nSortCol;
	ListView_SortItemsEx(GetSafeHwnd(), SortCompare, this);
}

// Sort compare function.
// Use GetItem(LVITEM*) and text buffers with fixed size here for performance reasons.
/*static*/ int CALLBACK CFileList::SortCompare(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	int nRet = 0;
	const CFileList *pThis = reinterpret_cast<const CFileList*>(lParamSort);
	TCHAR lpszItem1[SORT_BUF_SIZE], lpszItem2[SORT_BUF_SIZE];

	LVITEM lvi1, lvi2;
	lvi1.iItem = static_cast<int>(lParam1);
	lvi2.iItem = static_cast<int>(lParam2);
	if (pThis->m_nSortCol >= 0)
	{
		lvi1.mask = lvi2.mask = LVIF_PARAM | LVIF_TEXT;
		lvi1.iSubItem = lvi2.iSubItem = pThis->m_nSortCol;
		lvi1.pszText = lpszItem1;
		lvi2.pszText = lpszItem2;
		lvi1.cchTextMax = lvi2.cchTextMax = SORT_BUF_SIZE;
	}
	else
	{
		lvi1.mask = lvi2.mask = LVIF_PARAM;
		lvi1.iSubItem = lvi2.iSubItem = 0;
	}

	pThis->GetItem(&lvi1);
	pThis->GetItem(&lvi2);

	switch (pThis->m_nSortCol)
	{
	case FileColumns::NameColumn :
		break;										// Sorting is performed below
	case FileColumns::SizeColumn :
		// With folders, the size column is empty.
		// To differentiate between empty files and folders, comparison uses file size + 1.
		nRet = (lpszItem1[0] ? 1 + _tstol(lpszItem1) : 0) - (lpszItem2[0] ? 1 + _tstol(lpszItem2) : 0);
		break;
	case FileColumns::TypeColumn :
	case FileColumns::OwnerColumn :
		// Use case insensitive collating comparison with text columns.
		// This requires that the correct locale has been set using setlocale().
		nRet = _tcsicoll(lpszItem1, lpszItem2);
	default :
		if (pThis->m_nSortCol < 0)
		{
			// No sort column specified. Sort by the index (the order when the list has been populated)
			//  by masking out the directory flag from the item data.
			nRet = static_cast<int>(lvi1.lParam & ~nDirectoryFlag) - static_cast<int>(lvi2.lParam & ~nDirectoryFlag);
		}
		else
		{
			// All other columns are sorted by case sensitive string comparison.
			// This requires that date columns use the ISO format.
			nRet = _tcscmp(lpszItem1, lpszItem2);
		}
	}
	// Sort by name with name column or when items are identical (e.g. same size or date).
	if (0 == nRet && pThis->m_nSortCol >= 0)
	{
		// Sort by file/directory name.
		// Folders are shown on top/bottom with ascending/descending order like within the Windows Explorer.
		if ((lvi1.lParam ^ lvi2.lParam) & nDirectoryFlag)
			nRet = (lvi1.lParam & nDirectoryFlag) ? -1 : 1;
		else										// Both items are normal files or folders
		{
			if (lvi1.iSubItem != FileColumns::NameColumn)
			{
				lvi1.mask = lvi2.mask = LVIF_TEXT;
				lvi1.iSubItem = lvi2.iSubItem = FileColumns::NameColumn;
				pThis->GetItem(&lvi1);
				pThis->GetItem(&lvi2);
			}
			// Use case insensitive collating comparison for the name.
			// This requires that the correct locale has been set using setlocale().
			nRet = _tcsicoll(lpszItem1, lpszItem2);
		}
	}
	return pThis->m_bSortDesc ? -nRet : nRet;
}

