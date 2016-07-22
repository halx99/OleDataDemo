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

// OleDataDemoDlg.cpp : implementation file


#include "stdafx.h"
#include "OleDataDemo.h"
#include "OleDataDemoDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// COleDataDemoDlg dialog



COleDataDemoDlg::COleDataDemoDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COleDataDemoDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void COleDataDemoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_List);
	DDX_Control(pDX, IDC_EDIT1, m_Edit);
	DDX_Control(pDX, IDC_CHECK_DRAG_HI, m_btnDragTextHi);
	DDX_Control(pDX, IDC_EDIT_SCALE, m_editScale);
	DDX_Control(pDX, IDC_RICHEDIT21, m_editRich);
	DDX_Control(pDX, IDC_INFO_LIST, m_InfoList);
	DDX_Control(pDX, IDC_INFO_STATIC, m_InfoTitle);
	DDX_Control(pDX, IDC_DI_TEXT_RADIO, m_DiTextBtn);
	DDX_Control(pDX, IDC_DI_WND_RADIO, m_DiWndBtn);
	DDX_Control(pDX, IDC_DI_FILE_RADIO, m_DiFileBtn);
	DDX_Control(pDX, IDC_DI_RES_RADIO, m_DiResBtn);
	DDX_Control(pDX, IDC_DI_SYS_RADIO, m_DiSysBtn);
	DDX_Control(pDX, IDC_DI_HWND_CHECK, m_DiSystemCheck);
	DDX_Control(pDX, IDC_DROP_DESC_CHECK, m_DropDescCheck);
	DDX_Control(pDX, IDC_PICT3_STATIC, m_Pict1);
	DDX_Control(pDX, IDC_PICT1_STATIC, m_Pict2);
	DDX_Control(pDX, IDC_CHECK_ALLCOLUMNS, m_checkAllColumns);
	DDX_Control(pDX, IDC_LISTVIEW_COMBO, m_comboListView);
}

BEGIN_MESSAGE_MAP(COleDataDemoDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_CBN_SELCHANGE(IDC_LISTVIEW_COMBO, OnCbnSelchangeListviewCombo)
	ON_EN_CHANGE(IDC_EDIT_SCALE, OnEnChangeEditScale)
	ON_BN_CLICKED(IDC_INFO_BTN, OnBnClickedInfoBtn)
	ON_BN_CLICKED(IDC_DI_TEXT_RADIO, OnBnClickedImageTypeRadio)
	ON_BN_CLICKED(IDC_DI_WND_RADIO, OnBnClickedImageTypeRadio)
	ON_BN_CLICKED(IDC_DI_FILE_RADIO, OnBnClickedImageTypeRadio)
	ON_BN_CLICKED(IDC_DI_RES_RADIO, OnBnClickedImageTypeRadio)
	ON_BN_CLICKED(IDC_DI_SYS_RADIO, OnBnClickedImageTypeRadio)
	ON_BN_CLICKED(IDC_CHECK_DRAG_HI, OnBnClickedImageTypeCheck)
	ON_BN_CLICKED(IDC_DI_HWND_CHECK, OnBnClickedImageTypeCheck)
	ON_BN_CLICKED(IDC_DROP_DESC_CHECK, OnBnClickedImageTypeCheck)
	ON_BN_CLICKED(IDC_CHECK_ALLCOLUMNS, OnBnClickedCheckAllcolumns)
	ON_MESSAGE(WM_USER_UPDATE_INFO, OnUpdateInfo)
END_MESSAGE_MAP()


// COleDataDemoDlg message handlers

BOOL COleDataDemoDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// Register this window as drop target using a COleDropTargetEx class member.
	// This is the main window of this application. By registering as drop target,
	//  drag images are shown when over this window and all client windows that does
	//  not handle drag images.
	VERIFY(m_DropTarget.Register(this));

	// Initialize controls. 
	// This includes enabling drop support which should be done in OnCreate() of 
	//  the controls but this is not called for template created windows.
	m_Pict1.Init(IDB_BITMAP_DRAGME);
	m_Pict2.Init(IDB_BITMAP_DRAGME);
	m_Edit.Init();
	m_Edit.SetAllowDropSelf(true);					// can copy/move by drag & drop inside edit control
	m_Edit.SetAllowDragMove(true);					// can move content by drag & drop

	int nItem = m_comboListView.AddString(_T("Detailed View"));
	m_comboListView.SetCurSel(nItem);
	m_comboListView.SetItemData(nItem, LV_VIEW_DETAILS);
	nItem = m_comboListView.AddString(_T("List View"));
	m_comboListView.SetItemData(nItem, LV_VIEW_LIST);
	nItem = m_comboListView.AddString(_T("Icon View"));
	m_comboListView.SetItemData(nItem, LV_VIEW_ICON);
	nItem = m_comboListView.AddString(_T("Small Icon View"));
	m_comboListView.SetItemData(nItem, LV_VIEW_SMALLICON);
	nItem = m_comboListView.AddString(_T("Tile View"));
	m_comboListView.SetItemData(nItem, LV_VIEW_TILE);

	OnCbnSelchangeListviewCombo();					// select report view for list
	m_List.Init();
	m_List.SetPath(CSIDL_PERSONAL);
//	m_List.SetPath(_T("C:\\"));
	m_InfoList.Init(this);

	m_editScale.SetWindowText(_T("100"));
	m_editScale.SetLimitText(5);

#if 0
	// Populate the list control with some data
	CString s;
#define COLS 5
	for (int j = 0; j < COLS; j++)
	{
		s.Format(_T("Col%d"), j+1);
		m_List.InsertColumn(j, s.GetString(), LVCFMT_LEFT, 60, j);
	}
	for (int i = 0; i < 20; i++)
	{
		s.Format(_T("Item %d-%d"), i+1, 1);
		m_List.InsertItem(i, s.GetString());
		m_List.SetItemData(i, i);
		for (int j = 0; j < COLS; j++)
		{
			s.Format(_T("Item %d-%d"), i+1, j+1);
			m_List.SetItemText(i, j, s.GetString());
		}
	}
#endif
	// Select drag image type
	m_DiTextBtn.SetCheck(BST_CHECKED);

#if WINVER < 0x0600
	OSVERSIONINFOEX VerInfo;
	DWORDLONG dwlConditionMask = 0;
	::ZeroMemory(&VerInfo, sizeof(OSVERSIONINFOEX));
	VerInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
	VerInfo.dwMajorVersion = 6;
	VER_SET_CONDITION(dwlConditionMask, VER_MAJORVERSION, VER_GREATER_EQUAL);
	BOOL bVistaOrLater = ::VerifyVersionInfo(&VerInfo, VER_MAJORVERSION, dwlConditionMask);
	BOOL bThemed = bVistaOrLater && ::IsAppThemed();
	if (bThemed)
		m_DropDescCheck.SetCheck(BST_CHECKED);
	else
	{
		m_DropDescCheck.EnableWindow(FALSE);
		if (bVistaOrLater)
			AfxMessageBox(_T("Visual styles are disabled for this application.\nDrag images will not be shown."));
	}
#else
	if (::IsAppThemed())
		m_DropDescCheck.SetCheck(BST_CHECKED);
	else
	{
		m_DropDescCheck.EnableWindow(FALSE);
		AfxMessageBox(_T("Visual styles are disabled for this application.\nDrag images will not be shown."));
	}
#endif

	// Enable/disable dependant controls
	SetDragImageType(true);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void COleDataDemoDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void COleDataDemoDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR COleDataDemoDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// Apply list view style when it has been changed
void COleDataDemoDlg::OnCbnSelchangeListviewCombo()
{
	int nSel = m_comboListView.GetCurSel();
	if (nSel >= 0)
		m_List.SetView(static_cast<DWORD>(m_comboListView.GetItemData(nSel)));
}

// Pass drag image scaling value to controls upon changings
void COleDataDemoDlg::OnEnChangeEditScale()
{
	CString str;
	m_editScale.GetWindowText(str);
	int nScale = _tstoi(str.GetString());
	m_Edit.SetDragImageScale(nScale);
	m_editScale.SetDragImageScale(nScale);
	m_List.SetDragImageScale(nScale);
	m_Pict1.SetDragImageScale(nScale);
	m_Pict2.SetDragImageScale(nScale);
}

// Toggle between all and automatic columns in OLE data list
void COleDataDemoDlg::OnBnClickedCheckAllcolumns()
{
	if (BST_CHECKED == m_checkAllColumns.GetCheck())
		m_InfoList.AllColumns();
	else
		m_InfoList.DefaultColumns();
}

// Reload info list with clipboard data
void COleDataDemoDlg::OnBnClickedInfoBtn()
{
	m_InfoList.UpdateFromClipboard();
}

// WM_USER_UPDATE_INFO handler.
// May be sent by controls to update the info list when dropping.
LRESULT COleDataDemoDlg::OnUpdateInfo(WPARAM wParam, LPARAM /*lParam*/)
{
	m_InfoList.UpdateFromData(reinterpret_cast<COleDataObject *>(wParam));
	return 0;
}

// Set Info list title. 
// Called by info list control when content has been updated with drag data by dragging over control.
void COleDataDemoDlg::SetInfoTitle(LPCTSTR lpszTitle)
{
	m_InfoTitle.SetWindowText(lpszTitle);
}


// Drag image option check box changed.
void COleDataDemoDlg::OnBnClickedImageTypeCheck()
{
	SetDragImageType(false);
}

// Type of drag image changed.
void COleDataDemoDlg::OnBnClickedImageTypeRadio()
{
	SetDragImageType(true);
}

// Pass selected drag image type options to controls and
//  optional enable/disable dependant controls.
void COleDataDemoDlg::SetDragImageType(bool bDependencies)
{
	// Allow drop description text
	bool bDropDesc = (BST_CHECKED == m_DropDescCheck.GetCheck());

	// Drag image type radio buttons
	int nImageType = DRAG_IMAGE_FROM_TEXT;
	if (BST_CHECKED == m_DiWndBtn.GetCheck())
	{
		nImageType = DRAG_IMAGE_FROM_CAPT;
	}
	else if (BST_CHECKED == m_DiFileBtn.GetCheck())
	{
		nImageType = DRAG_IMAGE_FROM_EXT;
	}
	else if (BST_CHECKED == m_DiResBtn.GetCheck())
	{
		nImageType = DRAG_IMAGE_FROM_RES;
	}
	else if (BST_CHECKED == m_DiSysBtn.GetCheck())
	{
		nImageType = DRAG_IMAGE_FROM_SYSTEM;
	}
	else // if (BST_CHECKED == m_DiTextBtn.GetCheck())
	{
		// Use highlight text colors check box
		if (BST_CHECKED == m_btnDragTextHi.GetCheck())
			nImageType |= DRAG_IMAGE_HITEXT;
	}

	// Use internal image creation with list and tree view controls 
	if (DRAG_IMAGE_FROM_SYSTEM == nImageType && BST_CHECKED == m_DiSystemCheck.GetCheck())
		m_List.SetDragImageType(DRAG_IMAGE_FROM_HWND, bDropDesc);
	else
		m_List.SetDragImageType(nImageType, bDropDesc);
	m_Edit.SetDragImageType(nImageType, bDropDesc);
	m_editScale.SetDragImageType(nImageType, bDropDesc);
	m_Pict1.SetDragImageType(nImageType, bDropDesc);
	m_Pict2.SetDragImageType(nImageType, bDropDesc);

	// See notes at CMyListCtrl::DelayRenderOleData()
	// Due to a bug? in COleDataSource::Lookup(), multiple virtual files
	//  can't be defined for delay rendering.
	m_List.SetDragRendering(false);
//	m_Pict1.SetDragRendering(false);
//	m_Pict2.SetDragRendering(false);

	if (bDependencies)
	{
		m_btnDragTextHi.EnableWindow(DRAG_IMAGE_FROM_TEXT == (nImageType & ~DRAG_IMAGE_HITEXT));
		m_editScale.EnableWindow(DRAG_IMAGE_FROM_CAPT == nImageType || DRAG_IMAGE_FROM_EXT == nImageType);
		m_DiSystemCheck.EnableWindow(DRAG_IMAGE_FROM_SYSTEM == nImageType);
	}
}
