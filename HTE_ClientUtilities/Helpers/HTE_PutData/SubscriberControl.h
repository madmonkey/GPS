/* --------------------------------------------------------
 * HTE_PubData.SubscriberControl Header
 * $Workfile: SubscriberControl.h $
 * 
 * Listens for publish data of the same Topic.
 * The CSubscriber object implements everything needed, but in order to sink events in IE,
 * you must be an ActiveX control (instead of a plain COM object).
 * This control just blind aggregates interfaces (IDispatch, ISubscriber and _ISubscriberEvents)
 * implemented in the CSubscriber COM object.
 *
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 4 $
 * $Author: Steve $
 * $Modtime: 8/03/99 6:07p $
 * --------------------------------------------------------
 * $History: SubscriberControl.h $ 
 * 
 * *****************  Version 4  *****************
 * User: Steve        Date: 8/03/99    Time: 6:10p
 * Updated in $/Utilities/Misc/HTE_PubData
 * SubscriberControl was not properly aggregating the contained
 * CSubscriber which cause a memory leak and a crash when firing e.vents
 * Fixed control to blind aggregate CSubscriber and also aggregate the
 * events ising AggCP.h.
 * 
 * *****************  Version 3  *****************
 * User: Steve        Date: 5/07/99    Time: 10:11a
 * Updated in $/Utilities/Misc/HTE_PubData
 * Release
 * 
 * *****************  Version 2  *****************
 * User: Steve        Date: 4/27/99    Time: 10:53a
 * Updated in $/Utilities/Misc/HTE_PubData
 * Added Component categories
 * SafeForScripting
 * 
 * *****************  Version 1  *****************
 * User: Steve        Date: 4/27/99    Time: 10:20a
 * Created in $/Utilities/Misc/HTE_PubData
 * --------------------------------------------------------
 */

// SubscriberControl.h : Declaration of the CSubscriberControl

#ifndef __SUBSCRIBERCONTROL_H_
#define __SUBSCRIBERCONTROL_H_

#include <objsafe.h>
#include <comdef.h>
#include "resource.h"       // main symbols
#include <atlctl.h>
#include "Subscriber.h"
#include "AggCP.h"

/////////////////////////////////////////////////////////////////////////////
// CSubscriberControl
//
// Interesting Notes:
// * There is no IDispatch implemented, this interface is Blind aggregated 
//		along with ISubscriber from the contained CSubscriber.
//		This means you won't see IDispatchImpl here
//
// * The _ISubscriberEvents on CSubscriber are also aggregated
//
// * We implement IObjectSafety so we can be used in IE without security complaints
//
class ATL_NO_VTABLE CSubscriberControl : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComControl<CSubscriberControl>,
	public IPersistStreamInitImpl<CSubscriberControl>,
	public IOleControlImpl<CSubscriberControl>,
	public IOleObjectImpl<CSubscriberControl>,
	public IOleInPlaceActiveObjectImpl<CSubscriberControl>,
	public IViewObjectExImpl<CSubscriberControl>,
	public IOleInPlaceObjectWindowlessImpl<CSubscriberControl>,
	public IConnectionPointContainerImpl<CSubscriberControl>,
	public IPersistStorageImpl<CSubscriberControl>,
	public ISpecifyPropertyPagesImpl<CSubscriberControl>,
	public IQuickActivateImpl<CSubscriberControl>,
	public IDataObjectImpl<CSubscriberControl>,
	public IProvideClassInfo2Impl<&CLSID_SubscriberControl, &DIID__ISubscriberEvents, &LIBID_HTE_PubData>,
	public IPropertyNotifySinkCP<CSubscriberControl>,
	public CComCoClass<CSubscriberControl, &CLSID_SubscriberControl>,
	public IObjectSafetyImpl<CSubscriberControl, INTERFACESAFE_FOR_UNTRUSTED_CALLER>,
	public IAggConnectionPointImpl<CSubscriberControl, &DIID__ISubscriberEvents>
{
private:
	LPUNKNOWN m_pInnerUnk;

public:
	CSubscriberControl() : IAggConnectionPointImpl<CSubscriberControl, &DIID__ISubscriberEvents>(&m_pInnerUnk)
	{
		m_pInnerUnk = NULL;
	}


DECLARE_REGISTRY_RESOURCEID(IDR_SUBSCRIBERCONTROL)
DECLARE_GET_CONTROLLING_UNKNOWN()
DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CSubscriberControl)
	COM_INTERFACE_ENTRY(IViewObjectEx)
	COM_INTERFACE_ENTRY(IViewObject2)
	COM_INTERFACE_ENTRY(IViewObject)
	COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
	COM_INTERFACE_ENTRY(IOleInPlaceObject)
	COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
	COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
	COM_INTERFACE_ENTRY(IOleControl)
	COM_INTERFACE_ENTRY(IOleObject)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
	COM_INTERFACE_ENTRY(IQuickActivate)
	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY(IDataObject)
	COM_INTERFACE_ENTRY(IProvideClassInfo)
	COM_INTERFACE_ENTRY(IProvideClassInfo2)
	COM_INTERFACE_ENTRY_AGGREGATE_BLIND(m_pInnerUnk)
	COM_INTERFACE_ENTRY(IObjectSafety)
END_COM_MAP()

BEGIN_PROP_MAP(CSubscriberControl)
	PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
	PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
	// Example entries
	// PROP_ENTRY("Property Description", dispid, clsid)
	// PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()

BEGIN_CONNECTION_POINT_MAP(CSubscriberControl)
	CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
	CONNECTION_POINT_ENTRY(DIID__ISubscriberEvents)
END_CONNECTION_POINT_MAP()

BEGIN_CATEGORY_MAP(CPublisher)
  IMPLEMENTED_CATEGORY(CATID_SafeForScripting)
  IMPLEMENTED_CATEGORY(CATID_SafeForInitializing)
END_CATEGORY_MAP()

BEGIN_MSG_MAP(CSubscriberControl)
	CHAIN_MSG_MAP(CComControl<CSubscriberControl>)
	DEFAULT_REFLECTION_HANDLER()
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);


	HRESULT FinalConstruct();
	void FinalRelease();


// IViewObjectEx
	DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

// ISubscriberControl
public:

	HRESULT OnDraw(ATL_DRAWINFO& di)
	{
		RECT& rc = *(RECT*)di.prcBounds;
		Rectangle(di.hdcDraw, rc.left, rc.top, rc.right, rc.bottom);

		SetTextAlign(di.hdcDraw, TA_CENTER|TA_BASELINE);
		LPCTSTR pszText = _T("HTE_PubData.SubscriberControl");
		TextOut(di.hdcDraw, 
			(rc.left + rc.right) / 2, 
			(rc.top + rc.bottom) / 2, 
			pszText, 
			lstrlen(pszText));

		return S_OK;
	}

};

#endif //__SUBSCRIBERCONTROL_H_
