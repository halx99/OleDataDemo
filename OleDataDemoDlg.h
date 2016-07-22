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

// OleDataDemoDlg.h : header file
//

#pragma once
#include "OleDropTargetEx.h"
#include "MyEdit.h"
#include "MyListCtrl.h"
#include "FileList.h"
#include "MyStatic.h"
#include "InfoList.h"
#include "afxwin.h"
#include "afxcmn.h"

// COleDataDemoDlg dialog
class COleDataDemoDlg : public CDialog
{
// Construction
public:
	COleDataDemoDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_OLEDATADEMO_DIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON				m_hIcon;
	CFileList			m_List;
//	CMyListCtrl			m_List;
	CMyEdit				m_Edit;
	CRichEditCtrl		m_editRich;
	COleDropTargetEx	m_DropTarget;
	CMyEdit				m_editScale;
	CButton				m_btnDragTextHi;
	CInfoList			m_InfoList;
	CStatic				m_InfoTitle;
	CMyStatic			m_Pict1;
	CMyStatic			m_Pict2;
	CComboBox			m_comboListView;
	CButton				m_DiTextBtn;
	CButton				m_DiWndBtn;
	CButton				m_DiFileBtn;
	CButton				m_DiResBtn;
	CButton				m_DiSysBtn;
	CButton				m_DiSystemCheck;
	CButton				m_DropDescCheck;
	CButton				m_checkAllColumns;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnEnChangeEditScale();
	afx_msg void OnBnClickedInfoBtn();
	afx_msg void OnBnClickedImageTypeRadio();
	afx_msg void OnBnClickedImageTypeCheck();
	afx_msg void OnBnClickedCheckAllcolumns();
	afx_msg void OnCbnSelchangeListviewCombo();
	LRESULT OnUpdateInfo(WPARAM wParam, LPARAM lParam);

	void SetDragImageType(bool bDependencies);

public:
	void SetInfoTitle(LPCTSTR lpszTitle);
	DECLARE_MESSAGE_MAP()
};
