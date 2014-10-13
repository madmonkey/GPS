/* --------------------------------------------------------
 * HTE_PubData.Publisher
 * $Workfile: Publisher.h $
 * 
 * Publishes data synchronously to Subscribers of the same topic.
 *
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 4 $
 * $Author: Steve $
 * $Modtime: 6/28/99 2:27p $
 * --------------------------------------------------------
 * $History: Publisher.h $ 
 * 
 * *****************  Version 4  *****************
 * User: Steve        Date: 6/28/99    Time: 6:18p
 * Updated in $/Utilities/Misc/HTE_PubData
 * Added Timeout property to interface (defaults to 5 sec.)
 * Changed SendMessage to SendMessageTimeout to avoid hung applications
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

#ifndef __PUBLISHER_H_
#define __PUBLISHER_H_

#include <objsafe.h>
#include <comdef.h>
#include "resource.h"       // main symbols
#include "RegisteredMsgs.h"

#include <list>

typedef std::list<HWND> ListHWND;

/////////////////////////////////////////////////////////////////////////////
// CPublisher
class ATL_NO_VTABLE CPublisher : 
	public CWindowImpl<CPublisher>,
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CPublisher, &CLSID_Publisher>,
	public ISupportErrorInfo,
	public IObjectSafetyImpl<CPublisher, INTERFACESAFE_FOR_UNTRUSTED_CALLER>,
	public IDispatchImpl<IPublisher, &IID_IPublisher, &LIBID_HTE_PubData>
{
private:
	HWND m_hWnd;
	RegisteredMessages msgs;
	ListHWND m_SubscriberList;
	bstr_t m_bsTopic;
	long m_lTimeout;

public:
	CPublisher() : m_bsTopic(_T("")), m_lTimeout(5000)
	{
	}

	HRESULT FinalConstruct()
	{
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
			lResult = ::SendMessageTimeout( HWND_BROADCAST, msgs.ON_CREATE_PUBLISHER, (UINT)m_hWnd, 0, SMTO_NORMAL, 500, &dwResult );

			if (dwResult || !lResult)
				ATLTRACE(_T("SendMessageTimeout Error. lResult=%d, dwResult=%d\n"), lResult, dwResult);
		}


		return hr;
	}

	void FinalRelease()
	{
		::SendMessage( m_hWnd, WM_CLOSE, 0, 0 );
	}

DECLARE_WND_CLASS(_T("HTE_PubData_Publisher")) 

DECLARE_REGISTRY_RESOURCEID(IDR_PUBLISHER)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CPublisher)
	COM_INTERFACE_ENTRY(IPublisher)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IObjectSafety)
END_COM_MAP()


BEGIN_MSG_MAP(CPublisher)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
	MESSAGE_RANGE_HANDLER( 0xC000, 0xFFFF, OnRegisteredMessage )
END_MSG_MAP()

BEGIN_CATEGORY_MAP(CPublisher)
  IMPLEMENTED_CATEGORY(CATID_SafeForScripting)
  IMPLEMENTED_CATEGORY(CATID_SafeForInitializing)
END_CATEGORY_MAP()


// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IPublisher
public:
	STDMETHOD(get_Topic)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Topic)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_Timeout)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_Timeout)(/*[in]*/ long newVal);
	STDMETHOD(SendString)(/*[in]*/ BSTR Tag, /*[in]*/ BSTR Data);
	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		::DestroyWindow( m_hWnd );
		return 0;
	}

	LRESULT OnRegisteredMessage(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

};

#endif //__PUBLISHER_H_
