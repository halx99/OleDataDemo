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

// OleDataDemo.cpp : Defines the class behaviors for the application.


#include "stdafx.h"
#include "OleDataDemo.h"
#include "OleDataDemoDlg.h"

#include <ObjIdl.h>		// IGlobal interface
#include <CGuid.h>		// CLSID_GlobalOptions

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// IGlobal interface (ObjIdl.h)
// Interface definition for older SDK / Visual Studio versions
//#if !defined(NTDDI_VERSION) || (NTDDI_VERSION < NTDDI_VISTA)
#ifndef __IGlobalOptions_INTERFACE_DEFINED__

typedef enum tagGLOBALOPT_PROPERTIES
{
    COMGLB_EXCEPTION_HANDLING     = 1,
    COMGLB_APPID                  = 2,
    COMGLB_RPC_THREADPOOL_SETTING = 3
} GLOBALOPT_PROPERTIES;

typedef enum tagGLOBALOPT_EH_VALUES
{
    COMGLB_EXCEPTION_HANDLE               = 0,
    COMGLB_EXCEPTION_DONOT_HANDLE_FATAL   = 1,
    COMGLB_EXCEPTION_DONOT_HANDLE         = COMGLB_EXCEPTION_DONOT_HANDLE_FATAL,
    COMGLB_EXCEPTION_DONOT_HANDLE_ANY     = 2
} GLOBALOPT_EH_VALUES;

#undef INTERFACE
#define INTERFACE   IGlobalOptions
DECLARE_INTERFACE_( IGlobalOptions, IUnknown ) 
{
	/* IUnknown methods */
	STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
	STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
	STDMETHOD_( ULONG, Release )( THIS ) PURE;

	/* IGlobalOptions methods */
	STDMETHOD( Set )( THIS_ GLOBALOPT_PROPERTIES, ULONG_PTR ) PURE;
	STDMETHOD( Query )( THIS_ GLOBALOPT_PROPERTIES, ULONG_PTR * ) PURE;
};

static const IID IID_IGlobalOptions =
{ 0x0000015B, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };

// This is declared in CGuid.h with newer SDK versions but without NTDDI_VERSION condition.
//extern const CLSID CLSID_GlobalOptions;
static const CLSID CLSID_GlobalOptions =
{ 0x0000034B, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };
#endif

// COleDataDemoApp

BEGIN_MESSAGE_MAP(COleDataDemoApp, CWinApp)
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()


// COleDataDemoApp construction

COleDataDemoApp::COleDataDemoApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only COleDataDemoApp object

COleDataDemoApp theApp;


// COleDataDemoApp initialization

BOOL COleDataDemoApp::InitInstance()
{
	// InitCommonControlsEx() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	INITCOMMONCONTROLSEX icc = {
		sizeof(icc),
//		0xFFFF
		ICC_WIN95_CLASSES | ICC_STANDARD_CLASSES
	};
	VERIFY(::InitCommonControlsEx(&icc));

	CWinApp::InitInstance();

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	// of your final executable, you should remove from the following
	// the specific initialization routines you do not need
	// Change the registry key under which our settings are stored
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization
//	SetRegistryKey(_T("JoeArndt"));

	// Initialize OLE support.
	VERIFY(AfxOleInit());

	VERIFY(AfxInitRichEdit2());

#if 1
	// Turn off COM exception handler.
	// See http://blogs.msdn.com/b/oldnewthing/archive/2011/01/20/10117963.aspx
	// I should have make notes here why I did this.
	// After two years I only remember that I had some problems with faulty code in COleDropSourceEx::GiveFeedback
	//  resulting in an exception that was not detected.
	OSVERSIONINFO VerInfo;
	VerInfo.dwOSVersionInfoSize = sizeof(VerInfo);
	::GetVersionEx(&VerInfo);
	if (VerInfo.dwMajorVersion >= 6)
	{
		IGlobalOptions *pGlobalOptions;
#if defined(IID_PPV_ARGS) // The IID_PPV_ARGS macro has been introduced with Visual Studio 2005
		HRESULT hr = ::CoCreateInstance(CLSID_GlobalOptions, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pGlobalOptions));
#else
		HRESULT hr = ::CoCreateInstance(CLSID_GlobalOptions, NULL, CLSCTX_INPROC_SERVER, 
			IID_IGlobalOptions, (LPVOID*)&pGlobalOptions);
#endif
		if (SUCCEEDED(hr))
		{
			// With Vista, CoInitializeSecurity() must be called before the IGlobalOptions interface can be used.
			// COMGLB_EXCEPTION_DONOT_HANDLE_ANY is not supported with Vista.
			// NOTE: This is not tested with Vista! It may be necessary to use other parameters!
			if (6 == VerInfo.dwMajorVersion && 0 == VerInfo.dwMinorVersion)
			{
				hr = ::CoInitializeSecurity(
					NULL,						// PSECURITY_DESCRIPTOR          pSecDesc,
					-1,							// LONG                          cAuthSvc,
					NULL,						// SOLE_AUTHENTICATION_SERVICE  *asAuthSvc,
					NULL,						// void                         *pReserved1,
					RPC_C_AUTHN_LEVEL_DEFAULT,	// DWORD                         dwAuthnLevel,
					RPC_C_IMP_LEVEL_IDENTIFY,	// DWORD                         dwImpLevel,
					NULL,						// void                         *pAuthList,
					EOAC_NONE,					// DWORD                         dwCapabilities,
					NULL						// void                         *pReserved3
				);
				if (SUCCEEDED(hr))
					hr = pGlobalOptions->Set(COMGLB_EXCEPTION_HANDLING, COMGLB_EXCEPTION_DONOT_HANDLE);
			}
			// With Windows 7, CoInitializeSecurity() must be only called for specific IGlobalOptions
			//  operations and parameters but not for those used here.
			else
				hr = pGlobalOptions->Set(COMGLB_EXCEPTION_HANDLING, COMGLB_EXCEPTION_DONOT_HANDLE_ANY);
			pGlobalOptions->Release();
		}
		TRACE1("Set GlobalOptions result: %#X\n", hr);
	}
#endif


	COleDataDemoDlg *pDlg = new COleDataDemoDlg;
	m_pMainWnd = pDlg;
	pDlg->DoModal();
	delete pDlg;

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
