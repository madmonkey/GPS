/* --------------------------------------------------------
 * IAggConnectionPointImpl
 * $Workfile: AggCP.h $
 * 
 * Copyright (c) 1999, Jason Whittington
 *
 * This code is released into the public domain.  
 * Use at your own risk.
 *
 * IAggConnectionPointImpl is designed to expose connection points 
 * from inner aggregates.  It does this by simply delegating the 
 * Advise and Unadvise calls to the inner.  To use this class derive
 * from it and initialize it's constructor with the IUnknown of the
 * inner.  (If the inner is a smart pointer you'll need to explicitly
 * pass the address of the .p member, e.g:
 *
 * class Foo : public IAggConnectionPointImpl<Foo, &IID_IInnerEvents)
 * {
 * 		Foo() : IAggConnectionPointImpl<Foo, &IID_IInnerEvents>(&m_pInner.p)
 *		...
 * }
 *
 * For more info see www.develop.com/hp/jasonw/aggcp.htm
 * --------------------------------------------------------
 * Last Modified: 
 * $Revision: 1 $
 * $Author: Steve $
 * $Modtime: 8/03/99 5:58p $
 * --------------------------------------------------------
 * $History: AggCP.h $ 
 * 
 * *****************  Version 1  *****************
 * User: Steve        Date: 8/03/99    Time: 5:58p
 * Created in $/Utilities/Misc/HTE_PubData
 * --------------------------------------------------------
 */

#ifndef __AGGCP_H
#define  __AGGCP_H

template <class T, const IID* pIID, class CDV = CComDynamicUnkArray >
class IAggConnectionPointImpl : public _ICPLocator<pIID>
{
public:
    IUnknown ** m_pInner;
	IAggConnectionPointImpl(IUnknown ** ppInner)
	{m_pInner = ppInner;}

	STDMETHOD(_LocCPQueryInterface)(REFIID riid, void ** ppvObject)
	{
		if (InlineIsEqualGUID(riid, IID_IConnectionPoint) || InlineIsEqualUnknown(riid))
		{
			if (ppvObject == NULL)
				return E_POINTER;
			*ppvObject = this;
			AddRef();
#ifdef _ATL_DEBUG_INTERFACES
			_Module.AddThunk((IUnknown**)ppvObject, _T("IConnectionPointImpl"), riid);
#endif // _ATL_DEBUG_INTERFACES
			return S_OK;
		}
		else
			return E_NOINTERFACE;
	}

	STDMETHOD(GetConnectionInterface)(IID* piid2)
	{
		if (piid2 == NULL)
			return E_POINTER;
		*piid2 = *pIID;
		return S_OK;
	}

	STDMETHOD(GetConnectionPointContainer)(IConnectionPointContainer** ppCPC)
	{
		T* pT = static_cast<T*>(this);
		// No need to check ppCPC for NULL since QI will do that for us
		return pT->QueryInterface(IID_IConnectionPointContainer, (void**)ppCPC);
	}
	
	STDMETHOD(Advise)(IUnknown* pUnkSink, DWORD* pdwCookie)
	{
		if(m_pInner == NULL) return CONNECT_E_CANNOTCONNECT;
		if(*m_pInner == NULL) return CONNECT_E_CANNOTCONNECT;
		return AtlAdvise(*m_pInner,pUnkSink, *pIID, pdwCookie);
	}
	
	STDMETHOD(Unadvise)(DWORD dwCookie)
	{
		if(m_pInner == NULL) return CONNECT_E_NOCONNECTION;

		return AtlUnadvise(*m_pInner, *pIID, dwCookie);
	}

	STDMETHOD(EnumConnections)(IEnumConnections** ppEnum)
	{
		if(*m_pInner == NULL) return E_FAIL;
		if(m_pInner == NULL) return E_FAIL;

		{
			CComPtr<IConnectionPointContainer> pCPC;
			CComPtr<IConnectionPoint> pCP;

			HRESULT hRes = (*m_pInner)->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
			if (SUCCEEDED(hRes))
				hRes = pCPC->FindConnectionPoint(*pIID, &pCP);
			if (SUCCEEDED(hRes))
				hRes = pCP->EnumConnections(ppEnum);
			return hRes;
		}
	}
};

#endif



