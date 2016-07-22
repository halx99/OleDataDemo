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

#include "DragDropHelper.h"

// COleDropSourceEx results
#define DRAG_RES_NOT_STARTED		0		// drag not started
#define DRAG_RES_STARTED			1		// drag started (timeout expired or mouse left the start rect)
#define DRAG_RES_CANCELLED			2		// cancelled (may be started or not)
#define DRAG_RES_RELEASED			4		// left mouse button released (may be started or not)
#define DRAG_RES_DROPPED			8		// data has been dropped (return effect not DROPEEFECT_NONE)
// Possible combinations:
// DRAG_RES_RELEASED: 
//  Drag has not been started because mouse button was released before leaving the start rect or the
//  time out expired. The button down handler that started the drag operation may simulate a single click
//  by passing the original message to the default handler and send a button up message afterwards.
// DRAG_RES_CANCELLED: 
//  Drag has not been started because another mouse button was activated or the ESC key has been
//  pressed before leaving the start rect or the time out expired.
// DRAG_RES_STARTED | DRAG_RES_CANCELLED: 
//  Drag has been started but whas cancelled by ESC or activating other mouse button.
// DRAG_RES_STARTED | DRAG_RES_RELEASED:
//  Drag has been finished by releasing the mouse button but the target did not drop
//  (return value is DROPEFFECT_NONE).
// DRAG_RES_STARTED | DRAG_RES_RELEASED | DRAG_RES_DROPPED:
//  Drag has been finished by releasing the mouse button and the target dropped the data
//  (return value is not DROPEFFECT_NONE).

// COleDropSource derived class to support DropDescription
class COleDropSourceEx : public COleDropSource
{
public:
	COleDropSourceEx(); 
#ifdef _DEBUG
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:
	bool			m_bSetCursor;					// internal flag set when Windows cursor must be set
	int				m_nResult;						// drag result flags
	LPDATAOBJECT	m_pIDataObj;					// set by COleDataSourceEx to its IDataObject
	const CDropDescription *m_pDropDescription;		// set by COleDataSourceEx to its CDropDescription member

	bool			SetDragImageCursor(DROPEFFECT dwEffect);

	virtual BOOL	OnBeginDrag(CWnd* pWnd);
	virtual SCODE	QueryContinueDrag(BOOL bEscapePressed, DWORD dwKeyState);
	virtual SCODE	GiveFeedback(DROPEFFECT dropEffect);

public:
	friend class COleDataSourceEx;
};

