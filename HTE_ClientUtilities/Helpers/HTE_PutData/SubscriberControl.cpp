/* --------------------------------------------------------
 * HTE_PubData.SubscriberControl Implementation
 * $Workfile: SubscriberControl.cpp $
 * 
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 2 $
 * $Author: Steve $
 * $Modtime: 8/03/99 5:58p $
 * --------------------------------------------------------
 * $History: SubscriberControl.cpp $ 
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
#include "SubscriberControl.h"

/////////////////////////////////////////////////////////////////////////////
// CSubscriberControl

HRESULT CSubscriberControl::FinalConstruct()
{
	ATLTRACE(_T("CSubscriberControl::FinalConstruct.\n"));

	// Create the inner aggregated object.
	HRESULT hr = CSubscriber::_CreatorClass::CreateInstance(
		GetControllingUnknown(), IID_IUnknown, (void**)&m_pInnerUnk );

	return hr;
}

void CSubscriberControl::FinalRelease()
{
	ATLTRACE(_T("CSubscriberControl::FinalRelease.\n"));

	if (m_pInnerUnk)
		m_pInnerUnk->Release();
}

