/* --------------------------------------------------------
 * HTE_PubData.Publisher Implementation
 * $Workfile: Publisher.cpp $
 * 
 * Publishes data synchronously to Subscribers of the same topic.
 *
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 3 $
 * $Author: Steve $
 * $Modtime: 7/29/99 7:44p $
 * --------------------------------------------------------
 * $History: Publisher.cpp $ 
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
 * User: Steve        Date: 6/28/99    Time: 6:18p
 * Updated in $/Utilities/Misc/HTE_PubData
 * Added Timeout property to interface (defaults to 5 sec.)
 * Changed SendMessage to SendMessageTimeout to avoid hung applications
 * 
 * *****************  Version 1  *****************
 * User: Steve        Date: 4/27/99    Time: 10:20a
 * Created in $/Utilities/Misc/HTE_PubData
 * --------------------------------------------------------
 */

#include "stdafx.h"
#include "HTE_PubData.h"
#include "Publisher.h"

/////////////////////////////////////////////////////////////////////////////
// CPublisher

STDMETHODIMP CPublisher::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IPublisher
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

LRESULT CPublisher::OnRegisteredMessage(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	if (uMsg == msgs.ON_CREATE_SUBSCRIBER)
	{
		ATLTRACE(_T("Publisher:ON_CREATE_SUBSCRIBER\n"));

		// A subscriber is alive, add it to my list
		m_SubscriberList.push_back( (HWND)wParam );
	}
	else if (uMsg == msgs.ON_DESTROY_SUBSCRIBER)
	{
		ATLTRACE(_T("Publisher:ON_DESTROY_SUBSCRIBER\n"));

		// A subscriber shut down, remove it from my list
		m_SubscriberList.remove( (HWND)wParam );
	}

	return 0;
}

STDMETHODIMP CPublisher::SendString(BSTR Tag, BSTR Data)
{
	ATLTRACE(_T("SendString. Start.\n"));
	//Data is send to dest window thru the WM_COPYDATA msg.

	//First create the structure for the data to be copied
	COPYDATASTRUCT cds;
	cds.dwData = 0;

	long cbTopicLen  = ::SysStringByteLen( (BSTR)m_bsTopic ) + 2;
	long cbTagLen  = ::SysStringByteLen( Tag ) + 2;
	long cbDataLen = ::SysStringByteLen( Data ) + 2;
	long lOffset;
	
	//alloc space for all three strings and three 4-byte lengths
	cds.cbData = sizeof(cbTopicLen) + cbTopicLen + sizeof(cbTagLen) + cbTagLen + sizeof(cbDataLen)+ cbDataLen;
	char* lpBuffer = new char[cds.cbData];
	cds.lpData = lpBuffer;

	//marshall the two strings into a single buffer
	lOffset = 0;
	::CopyMemory( lpBuffer + lOffset, &cbTopicLen, sizeof(cbTopicLen) );
	
	lOffset += sizeof(cbTopicLen);
	::CopyMemory( lpBuffer + lOffset, (BSTR)m_bsTopic, cbTopicLen );
	
	lOffset += cbTopicLen;
	::CopyMemory( lpBuffer + lOffset, &cbTagLen, sizeof(cbTagLen) );
	
	lOffset += sizeof(cbTagLen);
	::CopyMemory( lpBuffer + lOffset, Tag, cbTagLen );
	
	lOffset += cbTagLen;
	::CopyMemory( lpBuffer + lOffset, &cbDataLen, sizeof(cbDataLen) );

	lOffset += sizeof(cbDataLen);
	::CopyMemory( lpBuffer + lOffset, (void*)Data, cbDataLen );


	ATLTRACE(_T("SendString. List Size=%d\n"), m_SubscriberList.size());
	
	for (ListHWND::iterator i = m_SubscriberList.begin(); i != m_SubscriberList.end(); ++i)
	{
		ATLTRACE(_T("Sending message to HWND=0x%X\n"), (HWND)*i);
		DWORD dwResult;
		LRESULT lResult = ::SendMessageTimeout( (HWND)*i, WM_COPYDATA, (UINT)m_hWnd, (LPARAM)&cds, SMTO_ABORTIFHUNG | SMTO_NORMAL, m_lTimeout, &dwResult );
		DWORD dwLastErr = GetLastError();

		// if the call failed because of anything other than a timeout condition, 
		// then remove it from our subscriber list
		if (lResult == 0 && dwLastErr != 0 && dwLastErr != 1460)
			i = m_SubscriberList.erase( i );
	}

	delete cds.lpData;

	ATLTRACE(_T("SendString. End.\n"));

	return S_OK;
}

STDMETHODIMP CPublisher::get_Topic(BSTR *pVal)
{
	*pVal = m_bsTopic.copy();

	return S_OK;
}

STDMETHODIMP CPublisher::put_Topic(BSTR newVal)
{
	m_bsTopic = bstr_t( newVal, true );

	return S_OK;
}

STDMETHODIMP CPublisher::get_Timeout(long *pVal)
{
	*pVal = m_lTimeout;

	return S_OK;
}

STDMETHODIMP CPublisher::put_Timeout(long newVal)
{
	m_lTimeout = newVal;

	return S_OK;
}
