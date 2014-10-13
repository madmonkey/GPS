/* --------------------------------------------------------
 * HTE_PubData.Subscriber Implementation
 * $Workfile: Subscriber.cpp $
 * 
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 2 $
 * $Author: Steve $
 * $Modtime: 7/30/99 9:33a $
 * --------------------------------------------------------
 * $History: Subscriber.cpp $ 
 * 
 * *****************  Version 2  *****************
 * User: Steve        Date: 8/03/99    Time: 6:10p
 * Updated in $/Utilities/Misc/HTE_PubData
 * SubscriberControl was not properly aggregating the contained
 * CSubscriber which cause a memory leak and a crash when firing e.vents
 * Fixed control to blind aggregate CSubscriber and also aggregate the
 * events ising AggCP.h.
 * 
 * *****************  Version 1  *****************
 * User: Steve        Date: 4/27/99    Time: 10:20a
 * Created in $/Utilities/Misc/HTE_PubData
 * --------------------------------------------------------
 */

#include "stdafx.h"
#include "HTE_PubData.h"
#include "Subscriber.h"
#include <COMDEF.H>

/////////////////////////////////////////////////////////////////////////////
// CSubscriber

STDMETHODIMP CSubscriber::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ISubscriber
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}



LRESULT CSubscriber::OnRegisteredMessage(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	if (uMsg == msgs.ON_CREATE_PUBLISHER)
	{
		ATLTRACE(_T("Subscriber:ON_CREATE_PUBLISHER\n"));

		// if a publish is created after me, then notify just that publisher that I exist
		DWORD dwResult;
		::SendMessageTimeout( (HWND)wParam, msgs.ON_CREATE_SUBSCRIBER, (UINT)m_hWnd, 0, 0, MSG_TIMEOUT, &dwResult );
	}

	return 0;
}

LRESULT CSubscriber::OnCopyData(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	USES_CONVERSION;

	ATLTRACE(_T("CSubscriber::OnCopyData -> Start\n"));
	
	long cbTopicLen;
	long cbTagLen;
	long cbDataLen;
	long lOffset;
	char* lpBuffer;
	BSTR bstrTopic = NULL;
	BSTR bstrTag = NULL;
	BSTR bstrData = NULL;
	HWND hwndFrom = (HWND)wParam;
	COPYDATASTRUCT* pCds = (COPYDATASTRUCT*)lParam;

	_ASSERTE( hwndFrom );
	_ASSERTE( pCds );
	_ASSERTE( pCds->cbData >= sizeof(cbTopicLen) + sizeof(cbTagLen) + sizeof(cbDataLen) );


	// insture that we handle the message properly incase some program send us a copy dat when it shouldn't have
	if (pCds->cbData >= sizeof(cbTagLen) + sizeof(cbDataLen))
	{
		lpBuffer = (char*)pCds->lpData;

		//unmarshal out two strings
		//format: <topiclen><....topic....><taglen><....tagbuffer....><datalen><....databuffer....>
		lOffset = 0;
		::CopyMemory( &cbTopicLen, lpBuffer + lOffset, sizeof(cbTopicLen) );

		lOffset += sizeof(cbTopicLen);
		bstrTopic = ::SysAllocStringByteLen( lpBuffer + lOffset, cbTopicLen );

		//if this is for the same topic then continue with unpacking
		if (0 == wcscmp( bstrTopic, m_bsTopic))
		{
			lOffset += cbTopicLen;
			::CopyMemory( &cbTagLen, lpBuffer + lOffset, sizeof(cbTagLen) );

			lOffset += sizeof(cbTagLen);
			bstrTag = ::SysAllocStringByteLen( lpBuffer + lOffset, cbTagLen );

			lOffset += cbTagLen;
			::CopyMemory( &cbDataLen, lpBuffer + lOffset, sizeof(cbDataLen) );

			lOffset += sizeof(cbDataLen);
			bstrData = ::SysAllocStringByteLen( lpBuffer + lOffset, cbDataLen );

			Fire_OnReceiveString( bstrTag, bstrData );

			::SysFreeString( bstrTag );
			::SysFreeString( bstrData );
		}

		::SysFreeString( bstrTopic );
	}

	ATLTRACE(_T("CSubscriber::OnCopyData -> End\n"));

	bHandled = TRUE;
	return 0;
}

STDMETHODIMP CSubscriber::get_Topic(BSTR *pVal)
{
	*pVal = m_bsTopic.copy();

	return S_OK;
}

STDMETHODIMP CSubscriber::put_Topic(BSTR newVal)
{
	m_bsTopic = bstr_t( newVal, true );

	return S_OK;
}
