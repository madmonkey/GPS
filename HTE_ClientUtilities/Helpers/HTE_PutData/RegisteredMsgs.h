/* --------------------------------------------------------
 * HTE_PubData RegisteredMessages Header
 * $Workfile: RegisteredMsgs.h $
 * 
 * The registered messages code is shared between the publisher and subscriber.
 * Thus the code is puled out into a separate file.
 *
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 1 $
 * $Author: Steve $
 * $Modtime: 4/27/99 10:05a $
 * --------------------------------------------------------
 * $History: RegisteredMsgs.h $ 
 * 
 * *****************  Version 1  *****************
 * User: Steve        Date: 4/27/99    Time: 10:20a
 * Created in $/Utilities/Misc/HTE_PubData
 * --------------------------------------------------------
 */

#ifndef REGISTEREDMSGS_H__
#define REGISTEREDMSGS_H__

const UINT MSG_TIMEOUT = 500; 

typedef struct
{
	UINT ON_CREATE_PUBLISHER;
	UINT ON_CREATE_SUBSCRIBER;
	UINT ON_DESTROY_SUBSCRIBER;
} RegisteredMessages;

void RegisterMessages( RegisteredMessages& msgs );

#endif