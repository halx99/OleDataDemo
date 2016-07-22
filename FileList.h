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

#include "MyListCtrl.h"

// GetOwner() flags.
#define GET_OWNER_NO_DOMAIN			0				// don't prefix name with domain
#define GET_OWNER_DOMAIN			0x0001			// prefix name with domain according to additional flags
#define GET_OWNER_NO_WELLKNOWNGROUP	0x0002			// don't prefix with domain when well known group
#define GET_OWNER_NO_LOCAL			0x0004			// don't prefix with domain when local computer name
#define GET_OWNER_NO_BUILTIN		0x0008			// don't prefix with domain when builtin account
#define GET_OWNER_SIDSTR			0x8000			// return SID as string when getting the name fails

#define CACHED_GET_OWNER_DATA						// use cached values for computer name and bultin SID

// The CFileFind class does not provide public access to the WIN32_FIND_DATA member.
class CFileFindEx : public CFileFind
{
public:
	inline const WIN32_FIND_DATA* GetFindData() const
	{ return m_hContext ? (const WIN32_FIND_DATA*)m_pFoundInfo : NULL; }
	inline DWORD GetAttributes() const
	{ return m_hContext && m_pFoundInfo ? ((LPWIN32_FIND_DATA)m_pFoundInfo)->dwFileAttributes : 0; }
};

// CFileList

class CFileList : public CMyListCtrl
{
	DECLARE_DYNAMIC(CFileList)

public:
	CFileList();
	virtual ~CFileList();

	enum FileColumns
	{
		NameColumn = 0,
		ModifiedColumn,
//		AccessedColumn,
//		CreatedColumn,
		TypeColumn,
		SizeColumn,
		OwnerColumn,
		AttributeColumn
	};

	void SetColumns();
	int  SetPath(LPCTSTR lpszPath);
	int  SetPath(int nFolder);
	int  ParseDirectory();
	void Sort(int nSortCol, bool bAutoDirection);
	void GetPath(CString& strPath, int nItem) const;

	inline LPCTSTR GetFolderName() const { return m_strPath.GetString(); }
	inline int GetItemIndex(int nItem) const 
	{ return static_cast<int>(GetItemData(nItem) & ~nDirectoryFlag); }
	inline bool IsDirectory(int nItem) const 
	{ return nItem >= 0 && (GetItemData(nItem) & nDirectoryFlag); }

protected:
	static const DWORD nDirectoryFlag = (DWORD)(1 << (8 * sizeof(DWORD) - 1));
	bool		m_bSortDesc;
	int			m_nSortCol;
#ifdef CACHED_GET_OWNER_DATA
	TCHAR		m_lpszComputerName[MAX_COMPUTERNAME_LENGTH + 1];
	BYTE		m_pSidBuf[SECURITY_MAX_SID_SIZE];
#endif
	CString		m_strPath;
	CImageList	m_LargeImages;
	CImageList	m_SmallImages;

	LPCTSTR GetAttributeString(CString& str, DWORD dwAttr) const;
	int AddImage(CImageList& ImageList, HICON hIcon, int nInitial);
	BOOL IsLocalDomain(LPCTSTR lpszDomain);
	BOOL IsBuiltinSid(PSID pSid);
	bool GetOwner(CString& strOwner, LPCTSTR lpszFile, unsigned nFlags = (UINT)-1);

	static int CALLBACK SortCompare(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort);

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnLvnColumnclick(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclk(NMHDR *pNMHDR, LRESULT *pResult);
};


