#include "IEPrintWatermarkHelper.h"
#include <xpsprint.h>
#include <xpsobjectmodel.h>

/*
#if defined _M_IX86
  #import "..\..\Externals\Deviare\DeviareCOM.dll" raw_interfaces_only, named_guids, raw_dispinterfaces, auto_rename
#elif defined _M_X64
  #import "..\..\Externals\Deviare\DeviareCOM64.dll" raw_interfaces_only, named_guids, raw_dispinterfaces, auto_rename
#else
  #error Unsupported platform
#endif

using namespace Deviare2;

#if defined _M_IX86
  #define my_ssize_t long
  #define my_size_t unsigned long
#elif defined _M_X64
  #define my_ssize_t __int64
  #define my_size_t unsigned __int64
#endif
*/

//------------------------------------------------------------------------------

/*
static VOID AddWatermark(__in IXpsOMPageReference *lpPageRef);
*/

//------------------------------------------------------------------------------

BOOL APIENTRY DllMain(__in HMODULE hModule, __in DWORD ulReasonForCall, __in LPVOID lpReserved)
{
  switch (ulReasonForCall)
  {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
      break;
  }
  return TRUE;
}

//------------------------------------------------------------------------------

MY_EXPORT LONG WINAPI GetXpsAddresses(__in SIZE_T buffer)
{
  SIZE_T *lpAddresses = (SIZE_T*)buffer;
  CComPtr<IXpsOMObjectFactory> cXpsFactory;
  CComPtr<IXpsOMPageReference> cXpsPageRef;
  XPS_SIZE sPageSize;
  SIZE_T *lpVtbl;
#ifdef _DEBUG
  CHAR szBufA[1024];
#endif //_DEBUG
  HRESULT hRes;

#ifdef _DEBUG
  sprintf_s(szBufA, 1024, "IEPrintWatermarkhelper::GetXpsAddresses [Addr]: 0x%Ix", buffer);
  ::OutputDebugStringA(szBufA);
#endif //_DEBUG
  hRes = ::CoCreateInstance(__uuidof(XpsOMObjectFactory), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&cXpsFactory));
  if (SUCCEEDED(hRes))
  {
    sPageSize.width = 100.0;
    sPageSize.height = 100.0;
    hRes = cXpsFactory->CreatePageReference(&sPageSize, &cXpsPageRef);
    if (SUCCEEDED(hRes))
    {
      lpVtbl = *(SIZE_T**)((IXpsOMPageReference*)cXpsPageRef);
      lpAddresses[0] = lpVtbl[5]; //address of SetPage
#ifdef _DEBUG
      sprintf_s(szBufA, 1024, "IEPrintWatermarkhelper::GetXpsAddresses [Res]: 0x%Ix", lpAddresses[0]);
      ::OutputDebugStringA(szBufA);
#endif //_DEBUG
    }
  }
  return hRes;
}
