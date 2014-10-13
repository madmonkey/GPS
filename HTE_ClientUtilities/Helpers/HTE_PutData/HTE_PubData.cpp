/* --------------------------------------------------------
 * HTE_PubData
 * $Workfile: HTE_PubData.cpp $
 * 
 * This is the main module of the DLL which contains the DLL Exports.
 *
 * Note: Proxy/Stub Information
 *      To build a separate proxy/stub DLL, 
 *      run nmake -f HTE_PubDataps.mk in the project directory.
 *
 * Copyright © 1999 HTE-UCS, Inc.  All Rights Reserved.
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 2 $
 * $Author: Steve $
 * $Modtime: 8/02/99 3:56p $
 * --------------------------------------------------------
 * $History: HTE_PubData.cpp $ 
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
 * User: Steve        Date: 4/27/99    Time: 10:19a
 * Created in $/Utilities/Misc/HTE_PubData
 * --------------------------------------------------------
 */


#include "stdafx.h"
#include "resource.h"
#include <initguid.h>
#include "HTE_PubData.h"

#include "HTE_PubData_i.c"
#include "Publisher.h"
#include "Subscriber.h"
#include "SubscriberControl.h"


CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_Subscriber, CSubscriber)
OBJECT_ENTRY(CLSID_Publisher, CPublisher)
OBJECT_ENTRY(CLSID_SubscriberControl, CSubscriberControl)
END_OBJECT_MAP()

/////////////////////////////////////////////////////////////////////////////
// DLL Entry Point

extern "C"
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID /*lpReserved*/)
{
    if (dwReason == DLL_PROCESS_ATTACH)
    {
        _Module.Init(ObjectMap, hInstance, &LIBID_HTE_PubData);
        DisableThreadLibraryCalls(hInstance);
    }
    else if (dwReason == DLL_PROCESS_DETACH)
        _Module.Term();
    return TRUE;    // ok
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE

STDAPI DllCanUnloadNow(void)
{
    return (_Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    return _Module.RegisterServer(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
    return _Module.UnregisterServer(TRUE);
}


