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

// OleDataDemo.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

// Drag image selection
#define DRAG_IMAGE_FROM_CAPT		0		// create drag image by capturing from screen
#define DRAG_IMAGE_FROM_TEXT		1		// create drag image from text
#define DRAG_IMAGE_FROM_RES			2		// create drag image from bitmap resource
#define DRAG_IMAGE_FROM_HWND		3		// let control window create the drag image upon DI_GETDRAGIMAGE msg
#define DRAG_IMAGE_FROM_SYSTEM		4		// let Windows create a generic drag image
#define DRAG_IMAGE_FROM_EXT			5		// get drag image from file extension
#define DRAG_IMAGE_FROM_BITMAP		6		// let control window create the image
#define DRAG_IMAGE_FROM_MASK		0x7FFF
#define DRAG_IMAGE_HITEXT			0x8000	// use system highlight colors with drag images from text

#define WM_USER_UPDATE_INFO	(WM_USER + 1)

// COleDataDemoApp:
// See OleDataDemo.cpp for the implementation of this class
//

class COleDataDemoApp : public CWinApp
{
public:
	COleDataDemoApp();

// Overrides
	public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern COleDataDemoApp theApp;