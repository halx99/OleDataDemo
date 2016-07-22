//
// Copyright (c) 2015-2017 x-studio365 - All Rights Reserved.
//
//
// PropertiesViewBar.cpp: implementation of the CPropertiesWnd class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "x-studio.h"
#include "MainFrm.h"
#include "PropertiesWnd.h"
#include "CustomProperties.h"
#include "ExtendedPropCtrl.h"
#include "detail/VXPropBinder.h"
#include "resource.h"
#include "purelib/utils/nsconv.h"
#include "detail/ResourceManager.h"
#include "purelib/utils/singleton.h"

#include "ShellHelpers.h"
#include "DragDropHelpers.h"

using namespace purelib;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#define new DEBUG_NEW
#endif

/////////////////////////////////////////
// CPropListDropTarget
/////////////////////////////////////////////////////////////////////////////
// CBCGPPropList message handlers
class CPropListDropTarget : public CDragDropHelper
{
public:
	CPropListDropTarget() :
		_cRef(1), _psiDrop(NULL)
	{
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv)
	{
		static const QITAB qit[] =
		{
			QITABENT(CPropListDropTarget, IDropTarget),
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		long cRef = InterlockedDecrement(&_cRef);
		if (!cRef)
			delete this;
		return cRef;
	}

	// IDropTarget
	IFACEMETHODIMP DragOver(DWORD  grfKeyState, POINTL pt, DWORD *pdwEffect) override
	{
		CPoint point(pt.x, pt.y);
		_wndPropList->ScreenToClient(&point);

		auto prop = _wndPropList->HitTest(point);
		if (m_strDisplayName.IsEmpty() || prop == nullptr || !prop->IsAcceptDropFiles()) {
			*pdwEffect = DROPEFFECT_NONE;
			ChangeDropImageType(DROPIMAGE_NONE);
		}
		else {
			*pdwEffect = DROPEFFECT_COPY;
			ChangeDropImageType(DROPIMAGE_COPY);
		}

		return __super::DragOver(grfKeyState, pt, pdwEffect);
	}

	void ChangeDropImageType(DROPIMAGETYPE newType)
	{
		if (newType != _dropImageType) {
			_dropImageType = newType;
			if (_pdtobj != nullptr) {
				IShellItem *psi;
				HRESULT hr = CreateItemFromObject(_pdtobj, IID_PPV_ARGS(&psi));
				if (SUCCEEDED(hr))
				{
					PWSTR pszName;
					hr = psi->GetDisplayName(SIGDN_NORMALDISPLAY/*SIGDN_FILESYSPATH*/, &pszName);
					if (SUCCEEDED(hr))
					{
						SetDropTip(_pdtobj, _dropImageType, _pszDropTipTemplate ? _pszDropTipTemplate : L"%1", pszName);
						CoTaskMemFree(pszName);
					}
					psi->Release();
				}
				else {
					auto strName = GetDisplayFullNameFromObject(_pdtobj);
					if (!strName.IsEmpty()) {
						SetDropTip(_pdtobj, _dropImageType, _pszDropTipTemplate ? _pszDropTipTemplate : L"%1", strName);
					}
				}
			}
		}
	}

public:
	void Register(CBCGPPropList* wndPropList, DROPIMAGETYPE dropImageType, PCWSTR pszDropTipTemplate)
	{
		_wndPropList = wndPropList;
		this->InitializeDragDropHelper(wndPropList->GetSafeHwnd(), dropImageType, pszDropTipTemplate);
	}

	HRESULT OnDrop(IShellItemArray *psia, LPCTSTR lpszFullName, DWORD, POINTL pt) override
	{
		_FreeItem();

		if (psia != nullptr) {
			// hold the dropped item for later in _psiDrop
			HRESULT hr = GetItemAt(psia, 0, IID_PPV_ARGS(&_psiDrop));
			if (SUCCEEDED(hr))
			{
				CPoint point(pt.x, pt.y);
				_wndPropList->ScreenToClient(&point);

				auto pHit = _wndPropList->HitTest(point);
				if (pHit != nullptr && pHit->IsAcceptDropFiles()) {
					PWSTR pszName;
					if (SUCCEEDED(_psiDrop->GetDisplayName(SIGDN_FILESYSPATH, &pszName)))
					{
						auto oldVal = pHit->GetValue();
						pHit->SetValue(pszName);
						if (pHit->GetValue() != oldVal)
							_wndPropList->OnPropertyChanged(pHit);

						CoTaskMemFree(pszName);
					}
				}
			}
		}
		else { // Windows XP compatible
			CPoint point(pt.x, pt.y);
			_wndPropList->ScreenToClient(&point);

			auto pHit = _wndPropList->HitTest(point);
			if (pHit != nullptr && pHit->IsAcceptDropFiles()) {
				auto oldVal = pHit->GetValue();
				pHit->SetValue(lpszFullName);
				if (pHit->GetValue() != oldVal)
					_wndPropList->OnPropertyChanged(pHit);
			}
		}

		return S_OK;
	}

protected:
	void _FreeItem()
	{
		SafeRelease(&_psiDrop);
	}
private:
	long _cRef;
	CBCGPPropList* _wndPropList;
	IShellItem *_psiDrop;
};

/////////////////////////////////////////////////////////////////////////////
// CResourceViewBar

CPropertiesWnd::CPropertiesWnd()
{
	m_pOleDropTarget = nullptr;
    m_nComboHeight = 0;

    m_BorderColor = UXColor::Orange;
    m_FillBrush.SetColors(UXColor::LightSteelBlue, UXColor::White, UXBrush::BCGP_GRADIENT_RADIAL_TOP_LEFT, 0.75);
    m_TextBrush.SetColor(UXColor::SteelBlue);
}

CPropertiesWnd::~CPropertiesWnd()
{
	SafeRelease(&m_pOleDropTarget);
}

BEGIN_MESSAGE_MAP(CPropertiesWnd, CBorderedDockingBar)
    //{{AFX_MSG_MAP(CPropertiesWnd)
    ON_WM_CREATE()
    ON_WM_SIZE()
    ON_COMMAND(ID_SORTING_GROUPBYTYPE, OnSortingprop)
    ON_UPDATE_COMMAND_UI(ID_SORTING_GROUPBYTYPE, OnUpdateSortingprop)
    ON_COMMAND(ID_EXPAND_ALL, OnExpandAll)
    ON_UPDATE_COMMAND_UI(ID_EXPAND_ALL, OnUpdateExpandAll)
    ON_WM_SETFOCUS()
    ON_WM_SETTINGCHANGE()
    ON_WM_PAINT()
    //}}AFX_MSG_MAP
    ON_REGISTERED_MESSAGE(BCGM_PROPERTY_COMMAND_CLICKED, OnCommandClicked)
    ON_REGISTERED_MESSAGE(BCGM_PROPERTY_MENU_ITEM_SELECTED, OnMenuItemSelected)
    ON_REGISTERED_MESSAGE(BCGM_PROPERTY_GET_MENU_ITEM_STATE, OnGetMenuItemState)
    ON_REGISTERED_MESSAGE(BCGM_PROPERTY_CHANGED, OnPropertyChanged) // append
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CResourceViewBar message handlers

UXPropTreeItem* CPropertiesWnd::HitTestProp(POINT screenPoint)
{
	this->m_wndPropList.ScreenToClient(&screenPoint);

	auto pHit = this->m_wndPropList.HitTest(screenPoint);
	if (pHit != nullptr && pHit->IsAcceptDropFiles())
		return pHit;
	return nullptr;
}

void CPropertiesWnd::AdjustLayout()
{
    if (GetSafeHwnd() == NULL || (AfxGetMainWnd() != NULL && AfxGetMainWnd()->IsIconic()))
    {
        return;
    }

    CRect rectClient;
    GetClientRect(rectClient);

    int x = rectClient.left;
    int y = rectClient.top;
    int cx = rectClient.Width();
    int cy = rectClient.Height();

    m_wndObjectCombo.SetWindowPos(NULL, x + 1, y + 1, cx - 2, m_nComboHeight, SWP_NOACTIVATE | SWP_NOZORDER);

    m_wndPropList.SetWindowPos(NULL, x + 1, y + m_nComboHeight, cx - 2, cy - 1 - m_nComboHeight, SWP_NOACTIVATE | SWP_NOZORDER);
}

void CPropertiesWnd::UpdateObjectTypeName(const char* typeName)
{
    m_wndObjectCombo.DeleteString(0);
    m_wndObjectCombo.AddString(nsc::transcode(typeName).c_str());
    m_wndObjectCombo.SetCurSel(0);
}

void CPropertiesWnd::SetVSDotNetLook(BOOL bSet)
{
    m_wndPropList.SetVSDotNetLook(bSet);
    m_wndPropList.SetGroupNameFullWidth(bSet);
}

int CPropertiesWnd::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
    if (UXDockingControlBar::OnCreate(lpCreateStruct) == -1)
        return -1;

    CRect rectDummy;
    rectDummy.SetRectEmpty();

    // Create combo:
    const DWORD dwViewStyle = WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST
        /*| WS_BORDER | CBS_SORT | WS_CLIPSIBLINGS | WS_CLIPCHILDREN*/;

    // No Need combo list
    if (!m_wndObjectCombo.Create(dwViewStyle, rectDummy, this, 1))
    {
        TRACE0("Failed to create Properies Combo \n");
        return -1;      // fail to create
    }

    m_wndObjectCombo.m_bVisualManagerStyle = TRUE;

    m_wndObjectCombo.AddString (_T("ObjectType"));
    m_wndObjectCombo.SetCurSel (0);

    CRect rectCombo;
    m_wndObjectCombo.GetWindowRect(&rectCombo);
   

    if (!m_wndPropList.Create(WS_VISIBLE | WS_CHILD, rectDummy, this, 2))
    {
        TRACE0("Failed to create Properies Grid \n");
        return -1;      // fail to create
    }

    InitPropList();

    m_nComboHeight = rectCombo.Height();
    AdjustLayout();

	// Drag & Drop support, TODO: move to x-studio365
	m_pOleDropTarget = new CPropListDropTarget();
	m_pOleDropTarget->Register(&this->m_wndPropList, DROPIMAGE_NONE, L"%1");
    return 0;
}

void CPropertiesWnd::OnSize(UINT nType, int cx, int cy)
{
    __super::OnSize(nType, cx, cy);
    AdjustLayout();
}

void CPropertiesWnd::OnSortingprop()
{
    m_wndPropList.SetAlphabeticMode();
}

void CPropertiesWnd::OnUpdateSortingprop(CCmdUI* pCmdUI)
{
    pCmdUI->SetCheck(m_wndPropList.IsAlphabeticMode());
}

void CPropertiesWnd::OnExpandAll()
{
    m_wndPropList.SetAlphabeticMode(FALSE);
}

void CPropertiesWnd::OnUpdateExpandAll(CCmdUI* pCmdUI)
{
    pCmdUI->SetCheck(!m_wndPropList.IsAlphabeticMode());
}

void CPropertiesWnd::InitPropList()
{
    SetPropListFont();
    SetObjectComboFont();

    // Add commands:
    /*CStringList lstCommands;
    lstCommands.AddTail (_T("Add New"));
    lstCommands.AddTail (_T("Remove All"));

    m_wndPropList.SetCommands (lstCommands);*/

    // Add custom menu items:
    CStringList lstCustomMenuItem;

    lstCustomMenuItem.AddTail(purelib::nsc::transcode(resource_manager->translateWord("Refresh")).c_str());
    lstCustomMenuItem.AddTail(purelib::nsc::transcode(resource_manager->translateWord("Clear")).c_str());

    m_wndPropList.SetCustomMenuItems(lstCustomMenuItem);

    // Setup general look:
    m_wndPropList.EnableToolBar();
    m_wndPropList.EnableSearchBox();
    m_wndPropList.EnableHeaderCtrl(FALSE);
    m_wndPropList.EnableDesciptionArea();
    m_wndPropList.MarkModifiedProperties();
    m_wndPropList.EnableContextMenu();

    SetVSDotNetLook();
}

void CPropertiesWnd::OnSetFocus(CWnd* pOldWnd)
{
    UXDockingControlBar::OnSetFocus(pOldWnd);
    m_wndPropList.SetFocus();
}

void CPropertiesWnd::OnSettingChange(UINT uFlags, LPCTSTR lpszSection)
{
    UXDockingControlBar::OnSettingChange(uFlags, lpszSection);
    SetPropListFont();
    SetObjectComboFont();
}

void CPropertiesWnd::SetPropListFont()
{
    ::DeleteObject(m_fntPropList.Detach());

    LOGFONT lf;
    globalData.fontRegular.GetLogFont(&lf);

    NONCLIENTMETRICS info;
    info.cbSize = sizeof(info);

    globalData.GetNonClientMetrics(info);

    lf.lfHeight = info.lfMenuFont.lfHeight;
    lf.lfWeight = info.lfMenuFont.lfWeight;
    lf.lfItalic = info.lfMenuFont.lfItalic;

    m_fntPropList.CreateFontIndirect(&lf);

    m_wndPropList.SetFont(&m_fntPropList);
}

void CPropertiesWnd::SetObjectComboFont()
{
    ::DeleteObject(m_fntObjectCombo.Detach());

    LOGFONT lf;
    globalData.fontBold.GetLogFont(&lf);

    NONCLIENTMETRICS info;
    info.cbSize = sizeof(info);

    globalData.GetNonClientMetrics(info);

    lf.lfHeight = info.lfMenuFont.lfHeight;
    lf.lfWeight = info.lfMenuFont.lfWeight + 200;
    lf.lfItalic = info.lfMenuFont.lfItalic;

    m_fntObjectCombo.CreateFontIndirect(&lf);

    m_wndObjectCombo.SetFont(&m_fntObjectCombo);
}

LRESULT CPropertiesWnd::OnCommandClicked(WPARAM, LPARAM lp)
{
    CString str;

    switch (lp)
    {
    case 0:
        str = _T("Add new control dialog");
        break;

    case 1:
        str = _T("Remove all controls dialog");
        break;
    }

    halccMsgBox(str);
    SetFocus();
    return 0;
}

LRESULT CPropertiesWnd::OnPropertyChanged(WPARAM, LPARAM lParam)
{
    auto pProp = (UXPropTreeItem*)lParam;

    auto observer = reinterpret_cast<VXPropBinder*>(pProp->GetData());
    if (observer)
    {
        if (!pProp->IsKindOf(RUNTIME_CLASS(CButtonProp)) && !pProp->IsKindOf(RUNTIME_CLASS(CTwoButtonsProp))) {

            auto pFileProp = DYNAMIC_DOWNCAST(UXFileProp, pProp);
            if (pFileProp != nullptr && pFileProp->IsOpenFile()) 
            {
                auto value = pFileProp->GetValue();
                if (value.vt == VT_BSTR) {
                    auto translated = resource_manager->translate(nsc::transcode(pFileProp->GetValue().bstrVal));
                    pFileProp->SetValue(translated.c_str());
                }
            }

            observer->SetValue(pProp);
        }
    }

    return 0;
}

// no use, TODO:remove
LRESULT CPropertiesWnd::OnGetMenuItemState(WPARAM wp, LPARAM lp)
{
    int nMenuIndex = (int)wp;

    UXPropTreeItem* pProp = (UXPropTreeItem*)lp;
    ASSERT_VALID(pProp);

    UINT nState = 0;

    switch (nMenuIndex)
    {
    case 0:
    case 1:
        if (pProp->IsKindOf(RUNTIME_CLASS(UXFileProp)))
            nState &= ~MF_GRAYED;
        else
            nState |= MF_GRAYED;
        break;
#if 0 // unused
    case 2:    // Extract value to resource
        nState |= (MF_GRAYED | MF_CHECKED);
        break;

    case 3:    // Go to Value Definition
        nState |= MF_GRAYED;
        break;
#endif
    }

    return nState;
}

// no use, TODO:remove
LRESULT CPropertiesWnd::OnMenuItemSelected(WPARAM wp, LPARAM lp)
{
    int nMenuIndex = (int)wp;

    UXPropTreeItem* pProp = (UXPropTreeItem*)lp;
    ASSERT_VALID(pProp);

    if (nMenuIndex == 0) {
        pProp->m_nCustomFlags = 1;
        OnPropertyChanged(wp, lp);
    }
    else if (nMenuIndex == 1) { // Clear value: only for string currently.
        pProp->m_nCustomFlags = 2;
        pProp->SetValue(L"");
        OnPropertyChanged(wp, lp);
    }

    SetFocus();
    return 0;
}

void CPropertiesWnd::OnPaint()
{
    __super::OnPaint();
#if 0
    CPaintDC dc(this); // device context for painting

    CRect rectClient;
    GetClientRect(&rectClient);

    dc.FillRect(rectClient, &globalData.brBarFace);
#endif
}

