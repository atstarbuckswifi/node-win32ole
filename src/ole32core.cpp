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
  size_t u8len = strlen(u8s);
  size_t wclen = MultiByteToWideChar(CP_UTF8, 0, u8s, u8len, NULL, 0);
  wchar_t *wcs = (wchar_t *)malloc((wclen + 1) * sizeof(wchar_t));
  wclen = MultiByteToWideChar(CP_UTF8, 0, u8s, u8len, wcs, wclen + 1); // + 1
  wcs[wclen] = L'\0';
  return wcs; // ucs2 *** must be free later ***
}

char *wcs2mbs(const wchar_t *wcs)
{
  size_t mblen = WideCharToMultiByte(GetACP(), 0,
    (LPCWSTR)wcs, -1, NULL, 0, NULL, NULL);
  char *mbs = (char *)malloc((mblen + 1));
  mblen = WideCharToMultiByte(GetACP(), 0,
    (LPCWSTR)wcs, -1, mbs, mblen, NULL, NULL); // not + 1
  mbs[mblen] = '\0';
  return mbs; // locale mbs *** must be free later ***
}

char *wcs2u8s(const wchar_t *wcs)
{
  size_t mblen = WideCharToMultiByte(CP_UTF8, 0,
    (LPCWSTR)wcs, -1, NULL, 0, NULL, NULL);
  char *mbs = (char *)malloc((mblen + 1));
  mblen = WideCharToMultiByte(CP_UTF8, 0,
    (LPCWSTR)wcs, -1, mbs, mblen, NULL, NULL); // not + 1
  mbs[mblen] = '\0';
  return mbs; // locale mbs *** must be free later ***
}

// obsoleted functions

// locale mbs -> Unicode -> UTF-8 (without free when use _malloca)
string MBCS2UTF8(string mbs)
{
  if(mbs == "") return "";
  int wlen = MultiByteToWideChar(CP_ACP, 0, mbs.c_str(), -1, NULL, 0);
  WCHAR *wbuf = (WCHAR *)_malloca((wlen + 1) * sizeof(WCHAR));
  if(wbuf == NULL){
    throw "_malloca failed for MBCS to UNICODE";
    return "";
  }
  *wbuf = L'\0';
  if(MultiByteToWideChar(CP_ACP, 0, mbs.c_str(), -1, wbuf, wlen) <= 0){
    throw "can't convert MBCS to UNICODE";
    return "";
  }
  int ulen = WideCharToMultiByte(CP_UTF8, 0, wbuf, -1, NULL, 0, NULL, NULL);
  char *ubuf = (char *)_malloca((ulen + 1) * sizeof(char));
  if(ubuf == NULL){
    throw "_malloca failed for UNICODE to UTF8";
    return "";
  }
  *ubuf = '\0';
  if(WideCharToMultiByte(CP_UTF8, 0, wbuf, -1, ubuf, ulen, NULL, NULL) <= 0){
    throw "can't convert UNICODE to UTF8";
    return "";
  }
  ubuf[ulen] = '\0';
  return ubuf;
}

// UTF8 -> Unicode -> locale mbs (without free when use _malloca)
string UTF82MBCS(string utf8)
{
  if(utf8 == "") return "";
  int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, NULL, 0);
  WCHAR *wbuf = (WCHAR *)_malloca((wlen + 1) * sizeof(WCHAR));
  if(wbuf == NULL){
    throw "_malloca failed for UTF8 to UNICODE";
    return "";
  }
  *wbuf = L'\0';
  if(MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, wbuf, wlen) <= 0){
    throw "can't convert UTF8 to UNICODE";
    return "";
  }
  int slen = WideCharToMultiByte(CP_ACP, 0, wbuf, -1, NULL, 0, NULL, NULL);
  char *sbuf = (char *)_malloca((slen + 1) * sizeof(char));
  if(sbuf == NULL){
    throw "_malloca failed for UNICODE to MBCS";
    return "";
  }
  *sbuf = '\0';
  if(WideCharToMultiByte(CP_ACP, 0, wbuf, -1, sbuf, slen, NULL, NULL) <= 0){
    throw "can't convert UNICODE to MBCS";
    return "";
  }
  sbuf[slen] = '\0';
  return sbuf;
}

// locale mbs -> Unicode (allocate wcs, must free)
WCHAR *MBCS2WCS(string mbs)
{
  if(mbs == ""){
    throw "assigned empty string for MBCS to UNICODE";
    return NULL;
  }
  int wlen = MultiByteToWideChar(CP_ACP, 0, mbs.c_str(), -1, NULL, 0);
  WCHAR *wbuf = new WCHAR[wlen + 1];
  if(wbuf == NULL){
    throw "_malloca failed for MBCS to UNICODE";
    return NULL;
  }
  *wbuf = L'\0';
  if(MultiByteToWideChar(CP_ACP, 0, mbs.c_str(), -1, wbuf, wlen) <= 0){
    delete [] wbuf;
    throw "can't convert MBCS to UNICODE";
    return NULL;
  }
  return wbuf;
}

// Unicode -> locale mbs (not free wcs, without free when use _malloca)
string WCS2MBCS(WCHAR *wbuf)
{
  if(wbuf == NULL || wbuf[0] == L'\0') return "";
  int slen = WideCharToMultiByte(CP_ACP, 0, wbuf, -1, NULL, 0, NULL, NULL);
  char *sbuf = (char *)_malloca((slen + 1) * sizeof(char));
  if(sbuf == NULL){
    throw "_malloca failed for UNICODE to MBCS";
    return "";
  }
  *sbuf = '\0';
  if(WideCharToMultiByte(CP_ACP, 0, wbuf, -1, sbuf, slen, NULL, NULL) <= 0){
    throw "can't convert UNICODE to MBCS";
    return "";
  }
  sbuf[slen] = '\0';
  return sbuf;
}

// locale mbs -> BSTR (allocate bstr, must free)
BSTR MBCS2BSTR(string str)
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

string OLE32coreException::errorMessage(const char *m)
{
  ostringstream oss;
  oss << "OLE error: ";
  if(m) oss << "[" << m << "] ";
  oss << rmsg << endl;
  return oss.str();
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

OCVariant::OCVariant(string str)
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

void OCVariant::checkOLEresult(string msg)
{
  // bug ? comment (see old ole32core.cpp project)
  throw OLE32coreException(msg + "() failed\n");
}

// AutoWrap() - Automation helper function...
HRESULT OCVariant::AutoWrap(int autoType, VARIANT *pvResult, BSTR ptName, OCVariant **argchain, unsigned argLen)
{
  // bug ? comment (see old ole32core.cpp project)
  // execute at the first time to safety free argchain
  // Allocate memory for arguments...
  unsigned int size = argchain ? argLen : 0;
  VARIANT *pArgs = (VARIANT*)alloca(size * sizeof(VARIANT));
  for (unsigned int i = 0; i < size;  i++) {
    // bug ? comment (see old ole32core.cpp project)
    // will be reallocated BSTR whein using VariantCopy() (see by debugger)
    OCVariant *p = argchain[size - i - 1]; // arguments are passed in reverse order
    VariantInit(&pArgs[i]); // It will be free before copy.
    VariantCopy(&pArgs[i], &p->v);
    delete p;
  }
  // bug ? comment (see old ole32core.cpp project)
  // unexpected free original BSTR
  if (v.vt != VT_DISPATCH)
  {
    checkOLEresult("Called with a non-object. AutoWrap");
    return DISP_E_BADVARTYPE;
  }
  if(!v.pdispVal){
    checkOLEresult("Called with NULL IDispatch. AutoWrap");
    return E_POINTER;
  }
  // Convert down to ANSI (for error message only)
  char szName[256];
  WideCharToMultiByte(CP_ACP, 0,
    ptName, -1, szName, sizeof(szName), NULL, NULL);
  // Get DISPID for name passed...
  HRESULT hr = NULL;
  DISPID dispID;
  hr = v.pdispVal->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
  if(FAILED(hr)){
    ostringstream oss;
    oss << hr << " [" << szName << "] ";
    oss << "IDispatch::GetIDsOfNames AutoWrap";
    checkOLEresult(oss.str());
    return hr;
  }
  // Build DISPPARAMS
  DISPPARAMS dp = { NULL, NULL, 0, 0 };
  dp.cArgs = size;
  dp.rgvarg = pArgs;
  // Handle special-case for property-puts!
  DISPID dispidNamed = DISPID_PROPERTYPUT;
  if(autoType & DISPATCH_PROPERTYPUT){
    dp.cNamedArgs = 1;
    dp.rgdispidNamedArgs = &dispidNamed;
  }
  EXCEPINFO exceptInfo;
  // Make the call!
  hr = v.pdispVal->Invoke(dispID, IID_NULL,
    LOCALE_USER_DEFAULT, autoType, &dp, pvResult, &exceptInfo, NULL); // or _SYSTEM_ ?
  for (unsigned int i = 0; i < size; i++) {
    VariantClear(&pArgs[i]);
  }
  if(FAILED(hr)){
    ostringstream oss;
    oss << hr << " [" << szName << "] = [" << dispID << "]\r\n";
    if (hr == DISP_E_EXCEPTION)
    {
      // Convert down to ANSI (for error message only)
      if (exceptInfo.pfnDeferredFillIn) exceptInfo.pfnDeferredFillIn(&exceptInfo);
      if (!exceptInfo.bstrDescription && exceptInfo.scode) hr = exceptInfo.scode;
    }
    if (hr == DISP_E_EXCEPTION)
    {
      char szErrSource[256];
      int sourceLen = WideCharToMultiByte(CP_ACP, 0, exceptInfo.bstrSource, SysStringLen(exceptInfo.bstrSource), szErrSource, sizeof(szErrSource), NULL, NULL);
      char szErrDescription[256];
      int descLen = WideCharToMultiByte(CP_ACP, 0, exceptInfo.bstrDescription, SysStringLen(exceptInfo.bstrDescription), szErrDescription, sizeof(szErrDescription), NULL, NULL);
      szErrSource[sourceLen] = 0;
      szErrDescription[descLen] = 0;
      
      oss << szErrSource << ": ";
      oss << szErrDescription << "\r\n";
    } else {
      oss << errorFromCode(hr) << "\r\n";
    }
    oss << "(It always seems to be appeared at that time you mistake calling ";
    oss << "'obj.get { ocv->getProp() }' <-> 'obj.call { ocv->invoke() }'.) ";
    oss << "IDispatch::Invoke AutoWrap";
    checkOLEresult(oss.str());
    return hr;
  }
  return hr;
}

OCVariant *OCVariant::getProp(BSTR prop, OCVariant **argchain, unsigned argLen)
{
  OCVariant *r = new OCVariant();
  AutoWrap(DISPATCH_PROPERTYGET, &r->v, prop, argchain, argLen); // distinguish METHOD
  return r; // may be called with DISPATCH_PROPERTYGET|DISPATCH_METHOD
  // 'METHOD' may be called only with DISPATCH_PROPERTYGET
  // but 'PROPERTY' must not be called only with DISPATCH_METHOD
}

OCVariant *OCVariant::putProp(BSTR prop, OCVariant **argchain, unsigned argLen)
{
  AutoWrap(DISPATCH_PROPERTYPUT, NULL, prop, argchain, argLen);
  return this;
}

OCVariant *OCVariant::invoke(BSTR method, OCVariant **argchain, unsigned argLen, bool re)
{
  if(!re){
    AutoWrap(DISPATCH_METHOD, NULL, method, argchain, argLen);
    return this;
  }else{
    OCVariant *r = new OCVariant();
    AutoWrap(DISPATCH_METHOD | DISPATCH_PROPERTYGET, &r->v, method, argchain, argLen);
    return r; // may be called with DISPATCH_PROPERTYGET|DISPATCH_METHOD
    // 'METHOD' may be called only with DISPATCH_PROPERTYGET
    // but 'PROPERTY' must not be called only with DISPATCH_METHOD
  }
}

void OLE32core::checkOLEresult(string msg)
{
  throw OLE32coreException(msg + "() failed\n");
}

bool OLE32core::connect(string locale)
{
  if(!finalized) checkOLEresult("called twice ? connect");
  finalized = false;
  oldlocale = setlocale(LC_ALL, NULL);
  setlocale(LC_ALL, locale.c_str());
  CoInitialize(NULL); // Initialize COM for this thread...
  return true;
}

bool OLE32core::disconnect()
{
  if(finalized) checkOLEresult("called twice ? disconnect");
  finalized = true;
  CoUninitialize(); // Uninitialize COM for this thread...
#ifdef DEBUG
  setlocale(LC_ALL, "C");
#else
  setlocale(LC_ALL, oldlocale.c_str());
#endif
  return true;
}

} // namespace ole32core
