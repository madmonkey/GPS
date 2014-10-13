/* --------------------------------------------------------
 * HTE_PubData RegisteredMessages Implementation.
 * $Workfile: RegisteredMsgs.cpp $
 * 
 * The registered messages code is shared between the publisher and subscriber.
 * this registers window messages for inter-process communication.
 *
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 1 $
 * $Author: Steve $
 * $Modtime: 4/27/99 10:06a $
 * --------------------------------------------------------
 * $History: RegisteredMsgs.cpp $ 
 * 
 * *****************  Version 1  *****************
 * User: Steve        Date: 4/27/99    Time: 10:20a
 * Created in $/Utilities/Misc/HTE_PubData
 * --------------------------------------------------------
 */

#include <windows.h>
#include <tchar.h>
#include "RegisteredMsgs.h"

void RegisterMessages( RegisteredMessages& msgs )
{
	msgs.ON_CREATE_PUBLISHER = RegisterWindowMessage( _T("HTE_PUBDATA_ON_CREATE_PUBLISHER") );
	msgs.ON_CREATE_SUBSCRIBER = RegisterWindowMessage( _T("HTE_PUBDATA_ON_CREATE_SUBSCRIBER") );
	msgs.ON_DESTROY_SUBSCRIBER = RegisterWindowMessage( _T("HTE_PUBDATA_ON_DESTROY_SUBSCRIBER") );
}
