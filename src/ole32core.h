#ifndef __OLE32CORE_H__
#define __OLE32CORE_H__

#include <ole2.h>
#include <locale.h>

#include <vector>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>

#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <ctime>

namespace ole32core {

#define BDISPFUNCIN() do{std::cerr<<"-IN "<<__FUNCTION__<<std::endl;}while(0)
#define BDISPFUNCOUT() do{std::cerr<<"-OUT "<<__FUNCTION__<<std::endl;}while(0)
#define BDISPFUNCDAT(f, a, t) do{fprintf(stderr, (f), (a), (t));}while(0)
#if defined(_DEBUG) || defined(DEBUG)
#define DISPFUNCIN() BDISPFUNCIN()
#define DISPFUNCOUT() BDISPFUNCOUT()
#define DISPFUNCDAT(f, a, t) BDISPFUNCDAT((f), (a), (t))
#else // RELEASE
#define DISPFUNCIN() // BDISPFUNCIN()
#define DISPFUNCOUT() // BDISPFUNCOUT()
#define DISPFUNCDAT(f, a, t) // BDISPFUNCDAT((f), (a), (t))
#endif

#define BASSERT(x) chkerr(!!(x), __FILE__, __LINE__, __FUNCTION__, #x)
#define BVERIFY(x) BASSERT(x)
#if defined(_DEBUG) || defined(DEBUG)
#define DASSERT(x) BASSERT(x)
#define DVERIFY(x) DASSERT(x)
#else // RELEASE
#define DASSERT(x)
#define DVERIFY(x) (x)
#endif
extern std::string errorFromCode(DWORD code);
extern std::wstring errorFromCodeW(DWORD code);
extern BOOL chkerr(BOOL b, char *m, int n, char *f, char *e);

extern std::string to_s(int num);
extern wchar_t *u8s2wcs(const char *u8s); // UTF8 -> UCS2 (allocate wcs, must free)
extern char *wcs2mbs(const wchar_t *wcs); // UCS2 -> locale (allocate mbs, must free)
extern char *wcs2u8s(const wchar_t *wcs); // UCS2 -> UTF8 (allocate mbs, must free)

// obsoleted functions

// (allocate bstr, must free)
extern BSTR MBCS2BSTR(const std::string& str);
// (not free bstr)
extern std::string BSTR2MBCS(BSTR bstr);

struct ErrorInfo {
  WORD  wCode;
  std::wstring sSource;
  std::wstring sDescription;
  std::wstring sHelpFile;
  DWORD dwHelpContext;
  SCODE scode;
};

class OCVariant {
public:
  VARIANT v;
public:
  OCVariant(); // result
  OCVariant(const OCVariant &s); // copy
  OCVariant(bool c_boolVal); // VT_BOOL
  OCVariant(long lVal, VARTYPE type = VT_I4); // VT_I4
  OCVariant(double dblVal, VARTYPE type = VT_R8); // VT_R8
  OCVariant(BSTR bstrVal); // VT_BSTR (previous allocated)
  OCVariant(const std::string& str); // allocate and convert to VT_BSTR
  OCVariant(const wchar_t* str); // allocate and convert to VT_BSTR
  OCVariant(IDispatch* disp); // VT_DISPATCH
  OCVariant& operator=(const OCVariant& other);
  virtual ~OCVariant();
  void Clear();
};

class OCDispatch {
public:
  IDispatch* disp;
public:
  OCDispatch(); // result
  OCDispatch(IDispatch* d);
  OCDispatch(const OCDispatch &s); // copy
  OCDispatch& operator=(const OCDispatch& other);
  virtual ~OCDispatch();
  void Clear();
  ITypeInfo* getTypeInfo();
protected:
  ITypeInfo* info;
public:
  HRESULT invoke(WORD targetType, DISPID propID, VARIANT *pvResult, ErrorInfo& errorInfo, unsigned argLen, OCVariant **argchain);
};

class OLE32core {
protected:
  bool finalized;
  std::string oldlocale;
public:
  OLE32core() : finalized(true) {}
  virtual ~OLE32core() { if(!finalized) disconnect(); }
  HRESULT connect(const std::string& locale);
  HRESULT disconnect(void);
};

} // namespace ole32core

#endif // __OLE32CORE_H__
