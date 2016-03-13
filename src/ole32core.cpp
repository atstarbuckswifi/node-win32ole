/*
  ole32core.cpp
  This source is independent of node/v8.
*/

#include "ole32core.h"

using namespace std;

namespace ole32core {

BOOL chkerr(BOOL b, char *m, int n, char *f, char *e)
{
  if(b) return b;
  DWORD code = GetLastError();
  std::string msg = errorFromCode(code);
  fprintf(stderr, "ASSERT in %08x module %s(%d) @%s: %s\n", code, m, n, f, e);
  // fwprintf(stderr, L"error: %s", buf); // Some wchars can't see on console.
  fprintf(stderr, "error: %s", msg.c_str());
  return b;
}

std::string errorFromCode(DWORD code)
{
  WCHAR *buf;
  DWORD len = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER
      | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
    NULL, code, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
    (LPWSTR)&buf, 0, NULL);
  if (!len) return "";
  char *mbs = wcs2mbs(buf);
  LocalFree(buf);
  std::string msg(mbs);
  free(mbs);
  return msg;
}

std::wstring errorFromCodeW(DWORD code)
{
  WCHAR *buf;
  DWORD len = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER
    | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
    NULL, code, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
    (LPWSTR)&buf, 0, NULL);
  if (!len) return L"";
  std::wstring msg(buf, len);
  LocalFree(buf);
  return msg;
}

string to_s(int num)
{
  ostringstream ossnum;
  ossnum << num;
  return ossnum.str(); // create new string
}

wchar_t *u8s2wcs(const char *u8s)
{
  int u8len = (int)strlen(u8s);
  int wclen = MultiByteToWideChar(CP_UTF8, 0, u8s, u8len, NULL, 0);
  wchar_t *wcs = (wchar_t *)malloc((wclen + 1) * sizeof(wchar_t));
  wclen = MultiByteToWideChar(CP_UTF8, 0, u8s, u8len, wcs, wclen + 1); // + 1
  wcs[wclen] = L'\0';
  return wcs; // ucs2 *** must be free later ***
}

char *wcs2mbs(const wchar_t *wcs)
{
  int mblen = WideCharToMultiByte(GetACP(), 0,
    (LPCWSTR)wcs, -1, NULL, 0, NULL, NULL);
  char *mbs = (char *)malloc((mblen + 1));
  mblen = WideCharToMultiByte(GetACP(), 0,
    (LPCWSTR)wcs, -1, mbs, mblen, NULL, NULL); // not + 1
  mbs[mblen] = '\0';
  return mbs; // locale mbs *** must be free later ***
}

char *wcs2u8s(const wchar_t *wcs)
{
  int mblen = WideCharToMultiByte(CP_UTF8, 0,
    (LPCWSTR)wcs, -1, NULL, 0, NULL, NULL);
  char *mbs = (char *)malloc((mblen + 1));
  mblen = WideCharToMultiByte(CP_UTF8, 0,
    (LPCWSTR)wcs, -1, mbs, mblen, NULL, NULL); // not + 1
  mbs[mblen] = '\0';
  return mbs; // locale mbs *** must be free later ***
}

// obsoleted functions

// locale mbs -> BSTR (allocate bstr, must free)
BSTR MBCS2BSTR(const string& str)
{
  BSTR bstr;
  size_t len = str.length();
  WCHAR *wbuf = (WCHAR*)alloca((len + 1) * sizeof(WCHAR));
  mbstowcs(wbuf, str.c_str(), len);
  wbuf[len] = L'\0';
  bstr = ::SysAllocString(wbuf);
  return bstr;
}

// bug ? comment (see old project)
// BSTR -> locale mbs (not free bstr)
string BSTR2MBCS(BSTR bstr)
{
  string str;
  int len = ::SysStringLen(bstr);
  char *buf = (char*)alloca((len + 1) * sizeof(char));
  wcstombs(buf, bstr, len);
  buf[len] = '\0';
  str = buf;
  return str;
}

OCVariant::OCVariant()
{
  DISPFUNCIN();
  VariantInit(&v);
  DISPFUNCDAT("--construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant::OCVariant(const OCVariant &s)
{
  DISPFUNCIN();
  VariantInit(&v); // It will be free before copy.
  VariantCopy(&v, &s.v);
  DISPFUNCDAT("--copy construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant::OCVariant(bool c_boolVal)
{
  DISPFUNCIN();
  VariantInit(&v);
  v.vt = VT_BOOL;
  v.boolVal = c_boolVal ? VARIANT_TRUE : VARIANT_FALSE;
  DISPFUNCDAT("--construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant::OCVariant(IDispatch* disp)
{
  DISPFUNCIN();
  VariantInit(&v);
  v.vt = VT_DISPATCH;
  v.pdispVal = disp;
  if (disp) disp->AddRef();
  DISPFUNCDAT("--construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant::OCVariant(long lVal, VARTYPE type)
{
  DISPFUNCIN();
  VariantInit(&v);
  v.vt = type;
  v.lVal = lVal;
  DISPFUNCDAT("--construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant::OCVariant(double dblVal, VARTYPE type)
{
  DISPFUNCIN();
  VariantInit(&v);
  v.vt = VT_R8;
  v.dblVal = dblVal;
  DISPFUNCDAT("--construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant::OCVariant(BSTR bstrVal)
{
  DISPFUNCIN();
  VariantInit(&v);
  v.vt = VT_BSTR;
  v.bstrVal = bstrVal;
  DISPFUNCDAT("--construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant::OCVariant(const string& str)
{
  DISPFUNCIN();
  VariantInit(&v);
  v.vt = VT_BSTR;
  v.bstrVal = MBCS2BSTR(str);
  DISPFUNCDAT("--construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant::OCVariant(const wchar_t* str)
{
  DISPFUNCIN();
  VariantInit(&v);
  v.vt = VT_BSTR;
  v.bstrVal = ::SysAllocString(str);
  DISPFUNCDAT("--construction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
}

OCVariant& OCVariant::operator=(const OCVariant& other)
{
  DISPFUNCIN();
  VariantCopy(&v, (VARIANT *)&other.v);
  DISPFUNCDAT("--assignment-- %08p %08lx\n", &v, v.vt);
  DISPFUNCOUT();
  return *this;
}

OCVariant::~OCVariant()
{
  DISPFUNCIN();
  DISPFUNCDAT("--destruction-- %08p %08lx\n", &v, v.vt);
  DISPFUNCDAT("---(second step in)%d%d", 0, 0);
  VariantClear(&v); // need it
  DISPFUNCDAT("---(second step out)%d%d", 0, 0);
  DISPFUNCOUT();
}

void OCVariant::Clear()
{
  DISPFUNCIN();
  VariantClear(&v);
  DISPFUNCOUT();
}

OCDispatch::OCDispatch():disp(NULL), info(NULL)
{
  DISPFUNCIN();
  DISPFUNCDAT("--construction-- %08p\n", disp, 0);
  DISPFUNCOUT();
}

OCDispatch::OCDispatch(IDispatch* d):disp(d), info(NULL)
{
  DISPFUNCIN();
  if (disp) disp->AddRef();
  DISPFUNCDAT("--copy construction-- %08p\n", disp, 0);
  DISPFUNCOUT();
}

OCDispatch::OCDispatch(const OCDispatch &s) :disp(s.disp), info(s.info)
{
	DISPFUNCIN();
	if (disp) disp->AddRef();
	if (info) info->AddRef();
	DISPFUNCDAT("--copy construction-- %08p\n", disp, 0);
	DISPFUNCOUT();
}

OCDispatch& OCDispatch::operator=(const OCDispatch& other)
{
  DISPFUNCIN();
  if (disp != other.disp)
  {
    if (disp) disp->Release();
    disp = other.disp;
    if (disp) disp->AddRef();
  }
  if (info != other.info)
  {
    if (info) info->Release();
    info = other.info;
    if (info) info->AddRef();
  }
  DISPFUNCDAT("--assignment-- %08p\n", disp, 0);
  DISPFUNCOUT();
  return *this;
}

OCDispatch::~OCDispatch()
{
  DISPFUNCIN();
  DISPFUNCDAT("--destruction-- %08p\n", disp, 0);
  DISPFUNCDAT("---(second step in)%d%d", 0, 0);
  if (disp)
  {
    disp->Release();
    disp = NULL;
  }
  if (info)
  {
    info->Release();
    info = NULL;
  }
  DISPFUNCDAT("---(second step out)%d%d", 0, 0);
  DISPFUNCOUT();
}

void OCDispatch::Clear()
{
  DISPFUNCIN();
  if (disp)
  {
    disp->Release();
    disp = NULL;
  }
  if (info)
  {
    info->Release();
    info = NULL;
  }
  DISPFUNCOUT();
}

ITypeInfo* OCDispatch::getTypeInfo()
{
  if (!info && disp)
  {
    UINT typeCount;
    HRESULT hr = disp->GetTypeInfoCount(&typeCount);
    if (SUCCEEDED(hr) && typeCount)
    {
      hr = disp->GetTypeInfo(0, LOCALE_USER_DEFAULT, &info);
      if (FAILED(hr)) info = NULL;
    }
  }
  return info;
}

// AutoWrap() - Automation helper function...
HRESULT OCDispatch::invoke(WORD targetType, DISPID propID, VARIANT *pvResult, ErrorInfo& errorInfo, unsigned argLen, OCVariant **argchain)
{
  // bug ? comment (see old ole32core.cpp project)
  // unexpected free original BSTR
  if (!disp) {
    return E_POINTER;
  }
  // bug ? comment (see old ole32core.cpp project)
  // execute at the first time to safety free argchain
  // Allocate memory for arguments...
  unsigned int size = argchain ? argLen : 0;
  VARIANT *pArgs = size ? (VARIANT*)alloca(size * sizeof(VARIANT)) : NULL;
  for (unsigned int i = 0; i < size;  ++i) {
    // bug ? comment (see old ole32core.cpp project)
    // will be reallocated BSTR whein using VariantCopy() (see by debugger)
    OCVariant *p = argchain[size - i - 1]; // arguments are passed in reverse order
    VariantInit(&pArgs[i]); // It will be free before copy.
    VariantCopy(&pArgs[i], &p->v);
    delete p;
  }
  // Build DISPPARAMS
  DISPPARAMS dp = { NULL, NULL, 0, 0 };
  dp.cArgs = size;
  dp.rgvarg = pArgs;
  // Handle special-case for property-puts!
  DISPID dispidNamed = DISPID_PROPERTYPUT;
  if((targetType & DISPATCH_PROPERTYPUT) && size){
    if (pArgs[0].vt == VT_DISPATCH)
    { // PUTting DISPATCH values in must be done via "PROPERTYPUTREF", replace the bit being used
      targetType = ((targetType & ~DISPATCH_PROPERTYPUT) | DISPATCH_PROPERTYPUTREF);
    }
    dp.cNamedArgs = 1;
    dp.rgdispidNamedArgs = &dispidNamed;
  }
  EXCEPINFO exceptInfo;
  memset(&exceptInfo, sizeof(exceptInfo), 0);
  // Make the call!
  HRESULT hr;
  if (!info) getTypeInfo();
  if (info)
  {
    hr = DispInvoke(disp, info, propID, targetType, &dp, pvResult, &exceptInfo, NULL); // or _SYSTEM_ ?
  } else {
    hr = disp->Invoke(propID, IID_NULL, LOCALE_USER_DEFAULT, targetType, &dp, pvResult, &exceptInfo, NULL); // or _SYSTEM_ ?
  }
  for (unsigned int i = 0; i < size; ++i) {
    VariantClear(&pArgs[i]);
  }
  if (hr == DISP_E_EXCEPTION)
  {
    // cleanup the error message a bit
    if (exceptInfo.pfnDeferredFillIn) exceptInfo.pfnDeferredFillIn(&exceptInfo);
    errorInfo.wCode = exceptInfo.wCode;
    errorInfo.scode = exceptInfo.scode;
    errorInfo.dwHelpContext = exceptInfo.dwHelpContext;
    if (exceptInfo.bstrDescription)
    {
      errorInfo.sDescription = wstring(exceptInfo.bstrDescription, SysStringLen(exceptInfo.bstrDescription));
      SysFreeString(exceptInfo.bstrDescription);
    }
    if (exceptInfo.bstrSource)
    {
      errorInfo.sSource = wstring(exceptInfo.bstrSource, SysStringLen(exceptInfo.bstrSource));
      SysFreeString(exceptInfo.bstrSource);
    }
    if (exceptInfo.bstrHelpFile)
    {
      errorInfo.sHelpFile = wstring(exceptInfo.bstrHelpFile, SysStringLen(exceptInfo.bstrHelpFile));
      SysFreeString(exceptInfo.bstrHelpFile);
    }
    if (exceptInfo.scode) hr = HRESULT_FROM_WIN32(exceptInfo.scode);
  }
  return hr;
}

HRESULT OLE32core::connect(const string& locale)
{
  if(!finalized) return S_FALSE;
  oldlocale = setlocale(LC_ALL, NULL);
  setlocale(LC_ALL, locale.c_str());
  HRESULT hr = CoInitialize(NULL); // Initialize COM for this thread...
  if (SUCCEEDED(hr)) finalized = false;
  return hr;
}

HRESULT OLE32core::disconnect()
{
  if(finalized) return S_FALSE;
  finalized = true;
  CoUninitialize(); // Uninitialize COM for this thread...
#ifdef DEBUG
  setlocale(LC_ALL, "C");
#else
  setlocale(LC_ALL, oldlocale.c_str());
#endif
  return S_OK;
}

} // namespace ole32core
