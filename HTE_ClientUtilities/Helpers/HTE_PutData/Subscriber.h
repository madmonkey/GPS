/* --------------------------------------------------------
 * HTE_PubData.Subscriber Header
 * $Workfile: Subscriber.h $
 * 
 * Listens for publish data of the same Topic.
 *
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 3 $
 * $Author: Steve $
 * $Modtime: 8/03/99 5:21p $
 * --------------------------------------------------------
 * $History: Subscriber.h $ 
 * 
 * *****************  Version 3  *****************
 * User: Steve        Date: 8/03/99    Time: 6:10p
 * Updated in $/Utilities/Misc/HTE_PubData
 * SubscriberControl was not properly aggregating the contained
 * CSubscriber which cause a memory leak and a crash when firing e.vents
 * Fixed control to blind aggregate CSubscriber and also aggregate the
 * events ising AggCP.h.
 * 
 * *****************  Version 2  *****************
 * User: Steve        Date: 5/07/99    Time: 10:11a
 * Updated in $/Utilities/Misc/HTE_PubData
 * Release
 * 
 * *****************  Version 1  *****************
 * User: Steve        Date: 4/27/99    Time: 10:20a
 * Created in $/Utilities/Misc/HTE_PubData
 * --------------------------------------------------------
 */

// Subscriber.h : Declaration of the CSubscriber

#ifndef __SUBSCRIBER_H_
#define __SUBSCRIBER_H_

#include <comdef.h>
#include "resource.h"       // main symbols
#include "RegisteredMsgs.h"
#include "HTE_PubDataCP.h"

/////////////////////////////////////////////////////////////////////////////
// CSubscriber
class ATL_NO_VTABLE CSubscriber : 
	public CWindowImpl<CSubscriber>,
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSubscriber, &CLSID_Subscriber>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<CSubscriber>,
	public IDispatchImpl<ISubscriber, &IID_ISubscriber, &LIBID_HTE_PubData>,
	public CProxy_ISubscriberEvents< CSubscriber >
{
private:
	HWND m_hWnd;
	RegisteredMessages msgs;
	bstr_t m_bsTopic;

public:
	CSubscriber() : m_bsTopic(_T(""))
	{
	}

	HRESULT FinalConstruct()
	{
		ATLTRACE(_T("CSubscriber::FinalConstruct -> Start\n"));

		HRESULT hr = S_OK;

		//register our windows messages
		RegisterMessages( msgs );

		RECT rect;
		rect.bottom = 0;
		rect.top = 0;
		rect.left = 0;
		rect.right = 0;

		m_hWnd = Create( ::GetDesktopWindow(), rect );
		::ShowWindow( m_hWnd, SW_HIDE );

		if (!m_hWnd)
		{
			HRESULT_FROM_WIN32(GetLastError());
			hr = Error( _T("Window creation failed"), IID_IPublisher, hr );
		}
		else
		{
			DWORD dwResult;
			LRESULT lResult;
			lResult = ::SendMessageTimeout( HWND_BROADCAST, msgs.ON_CREATE_SUBSCRIBER, (UINT)m_hWnd, 0, SMTO_NORMAL, 500, &dwResult );

			if (dwResult || !lResult)
				ATLTRACE(_T("SendMessageTimeout Error. lResult=%d, dwResult=%d\n"), lResult, dwResult);
		}

		ATLTRACE(_T("CSubscriber::FinalConstruct -> End. HRESULT=0x%X\n"), hr);

		return hr;
	}

	void FinalRelease()
	{
		ATLTRACE(_T("CSubscriber::FinalRelease -> Start\n"));
		DWORD dwResult;
		::SendMessageTimeout( HWND_BROADCAST, msgs.ON_DESTROY_SUBSCRIBER, (UINT)m_hWnd, 0, SMTO_NORMAL, 500, &dwResult );
		::SendMessage( m_hWnd, WM_CLOSE, 0, 0 );
		ATLTRACE(_T("CSubscriber::FinalRelease -> End\n"));
	}

DECLARE_WND_CLASS(_T("HTE_PubData_Subscriber")) 

DECLARE_REGISTRY_RESOURCEID(IDR_SUBSCRIBER)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CSubscriber)
	COM_INTERFACE_ENTRY(ISubscriber)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CSubscriber)
CONNECTION_POINT_ENTRY(DIID__ISubscriberEvents)
END_CONNECTION_POINT_MAP()


BEGIN_MSG_MAP(CPublisher)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
	MESSAGE_RANGE_HANDLER( 0xC000, 0xFFFF, OnRegisteredMessage )
	MESSAGE_HANDLER(WM_COPYDATA, OnCopyData)
END_MSG_MAP()


// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ISubscriber
public:
	STDMETHOD(get_Topic)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Topic)(/*[in]*/ BSTR newVal);
	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		::DestroyWindow( m_hWnd );
		return 0;
	}

	LRESULT OnRegisteredMessage(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	LRESULT OnCopyData(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

};

#endif //__SUBSCRIBER_H_
