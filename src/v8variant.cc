/*
  v8variant.cc
*/

#include "v8variant.h"
#include <node.h>
#include <nan.h>

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

#define CHECK_OLE_ARGS(info, n, av0, av1) do{ \
    if(info.Length() < n) \
      return Nan::ThrowTypeError(__FUNCTION__ " takes exactly " #n " argument(s)"); \
    if(!info[0]->IsString()) \
      return Nan::ThrowTypeError(__FUNCTION__ " the first argument is not a Symbol"); \
    if(n != 1) av1 = info[1]; /* may not be Array */ \
    else if(info.Length() < 2) av1 = Nan::New<Array>(0); /* change none to Array[] */ \
    else if(info[1]->IsArray()) av1 = info[1]; /* Array */ \
    else return Nan::ThrowTypeError(__FUNCTION__ " the second argument is not an Array"); \
    av0 = info[0]; \
  }while(0)

Handle<Value> NewOleException(HRESULT hr)
{
  Handle<String> hMsg = Nan::New<String>((const uint16_t*)errorFromCodeW(hr).c_str()).ToLocalChecked();
  Local<v8::Value> args[2] = { Nan::New<Uint32>(hr), hMsg };
  Local<Object> target = Nan::New(module_target);
  Handle<Function> function = Handle<Function>::Cast(GET_PROP(target, "OLEException").ToLocalChecked());
  return Nan::NewInstance(function, 2, args).ToLocalChecked();
}

Handle<Value> NewOleException(HRESULT hr, const ErrorInfo& info)
{
  Handle<String> hMsg;
  if (info.sDescription.empty())
  {
    hMsg = Nan::New<String>((const uint16_t*)errorFromCodeW(hr).c_str()).ToLocalChecked();
  } else {
    hMsg = Nan::New<String>((const uint16_t*)info.sDescription.c_str()).ToLocalChecked();
  }
  if (info.scode)
  {
    hr = HRESULT_FROM_WIN32(info.scode);
  }
  else if (info.wCode)
  {
    hr = info.wCode; // likely a private error code
  }
  Local<v8::Value> args[2] = { Nan::New<Uint32>(hr), hMsg };
  Local<Object> target = Nan::New(module_target);
  Handle<Function> function = Handle<Function>::Cast(GET_PROP(target, "OLEException").ToLocalChecked());
  Handle<Object> errorObj = Nan::NewInstance(function, 2, args).ToLocalChecked();
  if (!info.sSource.empty())
  {
    Nan::Set(errorObj, Nan::New("source").ToLocalChecked(), Nan::New((const uint16_t*)info.sSource.c_str()).ToLocalChecked());
  }
  if (!info.sHelpFile.empty())
  {
    Nan::Set(errorObj, Nan::New("helpFile").ToLocalChecked(), Nan::New((const uint16_t*)info.sHelpFile.c_str()).ToLocalChecked());
  }
  if (info.dwHelpContext)
  {
    Nan::Set(errorObj, Nan::New("helpContext").ToLocalChecked(), Nan::New((uint32_t)info.dwHelpContext));
  }
  return errorObj;
}

Nan::Persistent<FunctionTemplate> V8Variant::clazz;

void V8Variant::Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target)
{
  Nan::HandleScope scope;
  Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(1);
  t->SetClassName(Nan::New("V8Variant").ToLocalChecked());
  Nan::SetPrototypeMethod(t, "isA", OLEIsA);
  Nan::SetPrototypeMethod(t, "vtName", OLEVTName);
  Nan::SetPrototypeMethod(t, "toBoolean", OLEBoolean);
  Nan::SetPrototypeMethod(t, "toInt32", OLEInt32);
  Nan::SetPrototypeMethod(t, "toNumber", OLENumber);
  Nan::SetPrototypeMethod(t, "toDate", OLEDate);
  Nan::SetPrototypeMethod(t, "toUtf8", OLEUtf8);
  Nan::SetPrototypeMethod(t, "toValue", OLEValue);
//  Nan::SetPrototypeMethod(t, "New", New);
  Nan::SetPrototypeMethod(t, "call", OLECall);
  Nan::SetPrototypeMethod(t, "get", OLEGet);
  Nan::SetPrototypeMethod(t, "set", OLESet);
/*
 In ParseUnaryExpression() < v8/src/parser.cc >
 v8::Object::ToBoolean() is called directly for unary operator '!'
 instead of v8::Object::valueOf()
 so NamedPropertyHandler will not be called
 Local<Boolean> ToBoolean(); // How to fake ? override v8::Value::ToBoolean
*/
  Local<ObjectTemplate> instancetpl = t->InstanceTemplate();
  Nan::SetCallAsFunctionHandler(instancetpl, OLECallComplete);
  Nan::SetNamedPropertyHandler(instancetpl, OLEGetAttr, OLESetAttr);
//  Nan::SetIndexedPropertyHandler(instancetpl, OLEGetIdxAttr, OLESetIdxAttr);
  Nan::SetPrototypeMethod(t, "Finalize", Finalize);
  Nan::Set(target, Nan::New("V8Variant").ToLocalChecked(), t->GetFunction());
  clazz.Reset(t);
}

OCVariant *V8Variant::CreateOCVariant(Handle<Value> v)
{
  if (v->IsNull() || v->IsUndefined()) {
    // todo: make separate undefined type
    return new OCVariant();
  }
  if (v.IsEmpty() || v->IsExternal() || v->IsNativeError() || v->IsFunction())
  {
    return NULL;
  }
// VT_USERDEFINED VT_VARIANT VT_BYREF VT_ARRAY more...
  if(v->IsBoolean() || v->IsBooleanObject()){
    return new OCVariant(Nan::To<bool>(v).FromJust());
  }else if(v->IsArray()){
// VT_BYREF VT_ARRAY VT_SAFEARRAY
    std::cerr << "[Array (not implemented now)]" << std::endl; return NULL;
    std::cerr.flush();
  }else if(v->IsInt32()){
    return new OCVariant((long)Nan::To<int32_t>(v).FromJust());
  }else if(v->IsUint32()){
    return new OCVariant((long)Nan::To<uint32_t>(v).FromJust(), VT_UI4);
#if(0) // may not be supported node.js / v8
  }else if(v->IsInt64()){
    return new OCVariant(Nan::To<int64_t>(v).FromJust());
#endif
  }else if(v->IsNumber() || v->IsNumberObject()){
    std::cerr << "[Number (VT_R8 or VT_I8 bug?)]" << std::endl;
    std::cerr.flush();
    return new OCVariant(Nan::To<double>(v).FromJust()); // double
  }else if(v->IsDate()){
    double d = Nan::To<double>(v).FromJust();
    time_t sec = (time_t)(d / 1000.0);
    int msec = (int)(d - sec * 1000.0);
    struct tm *t = localtime(&sec); // *** must check locale ***
    if(!t){
      std::cerr << "[Date may not be valid]" << std::endl;
      std::cerr.flush();
      return NULL;
    }
    SYSTEMTIME syst;
    syst.wYear = t->tm_year + 1900;
    syst.wMonth = t->tm_mon + 1;
    syst.wDay = t->tm_mday;
    syst.wHour = t->tm_hour;
    syst.wMinute = t->tm_min;
    syst.wSecond = t->tm_sec;
    syst.wMilliseconds = msec;
    SystemTimeToVariantTime(&syst, &d);
    return new OCVariant(d, VT_DATE); // date
  }else if(v->IsRegExp()){
    std::cerr << "[RegExp (bug?)]" << std::endl;
    std::cerr.flush();
    return new OCVariant((const wchar_t*)*String::Value(v->ToDetailString()));
  }else if(v->IsString() || v->IsStringObject()){
    return new OCVariant((const wchar_t*)*String::Value(v));
  }else if(v->IsObject()){
#if(0)
    std::cerr << "[Object (test)]" << std::endl;
    std::cerr.flush();
#endif
    V8Variant *v8v = V8Variant::Unwrap<V8Variant>(Nan::To<Object>(v).ToLocalChecked());
    if(!v8v){
      std::cerr << "[Object may not be valid (null V8Variant)]" << std::endl;
      std::cerr.flush();
      return NULL;
    }
    // std::cerr << ocv->v.vt;
    return new OCVariant(v8v->ocv);
  }else{
    std::cerr << "[unknown type (not implemented now)]" << std::endl;
    std::cerr.flush();
  }
  return NULL;
}

NAN_METHOD(V8Variant::OLEIsA)
{
  DISPFUNCIN();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  DISPFUNCOUT();
  return info.GetReturnValue().Set(Nan::New(v8v->ocv.v.vt));
}

NAN_METHOD(V8Variant::OLEVTName)
{
  DISPFUNCIN();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  Local<Object> target = Nan::New(module_target);
  Array *a = Array::Cast(*(GET_PROP(target, "vt_names").ToLocalChecked()));
  DISPFUNCOUT();
  return info.GetReturnValue().Set(ARRAY_AT(a, v8v->ocv.v.vt));
}

NAN_METHOD(V8Variant::OLEBoolean)
{
  DISPFUNCIN();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  if(v8v->ocv.v.vt != VT_BOOL)
    return Nan::ThrowTypeError("OLEBoolean source type OCVariant is not VT_BOOL");
  bool c_boolVal = v8v->ocv.v.boolVal != VARIANT_FALSE;
  DISPFUNCOUT();
  return info.GetReturnValue().Set(c_boolVal);
}

NAN_METHOD(V8Variant::OLEInt32)
{
  DISPFUNCIN();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  VARIANT& v = v8v->ocv.v;
  if(v.vt != VT_I4 && v.vt != VT_INT
  && v.vt != VT_UI4 && v.vt != VT_UINT)
    return Nan::ThrowTypeError("OLEInt32 source type OCVariant is not VT_I4 nor VT_INT nor VT_UI4 nor VT_UINT");
  DISPFUNCOUT();
  return info.GetReturnValue().Set(Nan::New(v.lVal));
}

NAN_METHOD(V8Variant::OLENumber)
{
  DISPFUNCIN();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  VARIANT& v = v8v->ocv.v;
  if(v.vt != VT_R8)
    return Nan::ThrowTypeError("OLENumber source type OCVariant is not VT_R8");
  DISPFUNCOUT();
  return info.GetReturnValue().Set(Nan::New(v.dblVal));
}

NAN_METHOD(V8Variant::OLEDate)
{
  DISPFUNCIN();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  VARIANT& v = v8v->ocv.v;
  if(v.vt != VT_DATE)
    return Nan::ThrowTypeError("OLEDate source type OCVariant is not VT_DATE");
  Local<Date> result = OLEDateToObject(v.date);
  DISPFUNCOUT();
  return info.GetReturnValue().Set(result);
}

NAN_METHOD(V8Variant::OLEUtf8)
{
  DISPFUNCIN();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  VARIANT& v = v8v->ocv.v;
  if(v.vt != VT_BSTR)
    return Nan::ThrowTypeError("OLEUtf8 source type OCVariant is not VT_BSTR");
  Handle<Value> result;
  if(!v.bstrVal) result = Nan::Undefined(); // or Null();
  else {
    std::wstring wstr(v.bstrVal);
    char *cs = wcs2u8s(wstr.c_str());
    result = Nan::New(cs).ToLocalChecked();
    free(cs);
  }
  DISPFUNCOUT();
  return info.GetReturnValue().Set(result);
}

NAN_METHOD(V8Variant::OLEValue)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OLETRACEVT(thisObject);
  OLETRACEFLUSH();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  if (!v8v) { std::cerr << "v8v is null"; std::cerr.flush(); }
  CHECK_V8V(v8v);
  VARIANT& v = v8v->ocv.v;
  switch(v.vt)
  {
  case VT_DISPATCH:
    if (v.pdispVal == NULL) {
      return info.GetReturnValue().SetNull();
    } else {
      return info.GetReturnValue().Set(thisObject); // through it
    }
  case VT_DISPATCH|VT_BYREF:
    if (!v.ppdispVal) return info.GetReturnValue().SetUndefined(); // really shouldn't happen
    if (*v.ppdispVal == NULL) {
      return info.GetReturnValue().SetNull();
    } else {
      return info.GetReturnValue().Set(thisObject); // through it
    }
  default:
    return info.GetReturnValue().Set(VariantToValue(info.This(), v));
  }
  OLETRACEOUT();
}

Local<Date> V8Variant::OLEDateToObject(const DATE& dt)
{
  DISPFUNCIN();
  SYSTEMTIME syst;
  VariantTimeToSystemTime(dt, &syst);
  struct tm t = { 0 }; // set t.tm_isdst = 0
  t.tm_year = syst.wYear - 1900;
  t.tm_mon = syst.wMonth - 1;
  t.tm_mday = syst.wDay;
  t.tm_hour = syst.wHour;
  t.tm_min = syst.wMinute;
  t.tm_sec = syst.wSecond;
  DISPFUNCOUT();
  return Nan::New<Date>(mktime(&t) * 1000.0 + syst.wMilliseconds).ToLocalChecked();
}

Local<Value> V8Variant::ArrayPrimitiveToValue(Handle<Object> thisObject, void* loc, VARTYPE vt, unsigned cbElements, unsigned idx)
{
  /*
  *  VT_CY               [V][T][P][S]  currency
  *  VT_UNKNOWN          [V][T]   [S]  IUnknown *
  *  VT_DECIMAL          [V][T]   [S]  16 byte fixed point
  *  VT_RECORD           [V]   [P][S]  user defined type
  */
  switch (vt)
  {
  case VT_DISPATCH:
    // ASSERT: cbElements == sizeof(IDispatch*)
    if (reinterpret_cast<IDispatch**>(loc)[idx] == NULL) {
      return Nan::Null();
    } else {
      Handle<Object> vResult = V8Variant::CreateUndefined();
      V8Variant *o = V8Variant::Unwrap<V8Variant>(vResult);
      CHECK_V8V_UNDEFINED(o);
      VARIANT& dest = o->ocv.v;
      dest.vt = VT_DISPATCH;
      dest.pdispVal = reinterpret_cast<IDispatch**>(loc)[idx];
      dest.pdispVal->AddRef();
      return vResult;
    }
  case VT_ERROR:
    // ASSERT: cbElements == sizeof(SCODE)
    return Exception::Error(Nan::New<String>((const uint16_t*)errorFromCodeW(HRESULT_FROM_WIN32(reinterpret_cast<SCODE*>(loc)[idx])).c_str()).ToLocalChecked());
  case VT_BOOL:
    // ASSERT: cbElements == sizeof(VARIANT_BOOL)
    return reinterpret_cast<VARIANT_BOOL*>(loc)[idx] != VARIANT_FALSE ? Nan::True() : Nan::False();
  case VT_I1:
    // ASSERT: cbElements == sizeof(CHAR)
    return Nan::New(reinterpret_cast<CHAR*>(loc)[idx]);
  case VT_UI1:
    // ASSERT: cbElements == sizeof(BYTE)
    return Nan::New(reinterpret_cast<BYTE*>(loc)[idx]);
  case VT_I2:
    // ASSERT: cbElements == sizeof(SHORT)
    return Nan::New(reinterpret_cast<SHORT*>(loc)[idx]);
  case VT_UI2:
    // ASSERT: cbElements == sizeof(USHORT)
    return Nan::New(reinterpret_cast<USHORT*>(loc)[idx]);
  case VT_I4:
    // ASSERT: cbElements == sizeof(LONG)
    return Nan::New(reinterpret_cast<LONG*>(loc)[idx]);
  case VT_UI4:
    // ASSERT: cbElements == sizeof(ULONG)
    return  Nan::New((uint32_t)reinterpret_cast<ULONG*>(loc)[idx]);
  case VT_INT:
    // ASSERT: cbElements == sizeof(INT)
    return Nan::New(reinterpret_cast<INT*>(loc)[idx]);
  case VT_UINT:
    // ASSERT: cbElements == sizeof(UINT)
    return Nan::New(reinterpret_cast<UINT*>(loc)[idx]);
  case VT_R4:
    // ASSERT: cbElements == sizeof(FLOAT)
    return Nan::New(reinterpret_cast<FLOAT*>(loc)[idx]);
  case VT_R8:
    // ASSERT: cbElements == sizeof(DOUBLE)
    return Nan::New(reinterpret_cast<DOUBLE*>(loc)[idx]);
  case VT_BSTR:
    // ASSERT: cbElements == sizeof(BSTR)
    if (!reinterpret_cast<BSTR*>(loc)[idx]) return Nan::Undefined(); // really shouldn't happen
    return Nan::New<String>((const uint16_t*)reinterpret_cast<BSTR*>(loc)[idx]).ToLocalChecked();
  case VT_DATE:
    // ASSERT: cbElements == sizeof(DATE)
    return OLEDateToObject(reinterpret_cast<DATE*>(loc)[idx]);
  case VT_VARIANT:
    // ASSERT: cbElements == sizeof(VARIANT)
    return VariantToValue(thisObject, reinterpret_cast<VARIANT*>(loc)[idx]);
  default:
    {
      Handle<Value> s = INSTANCE_CALL(thisObject, "vtName", 0, NULL);
      std::cerr << "[unknown type " << vt << ":" << *String::Utf8Value(s);
      std::cerr << " (not implemented now)]" << std::endl;
      std::cerr.flush();
    }
    return Nan::Undefined();
  }
}

Local<Value> V8Variant::ArrayToValue(Handle<Object> thisObject, const SAFEARRAY& a)
{
  OLETRACEIN();
  OLETRACEFLUSH();
  VARTYPE vt = VT_EMPTY;
  if (a.fFeatures&FADF_BSTR)
  {
    vt = VT_BSTR;
  }
  else if (a.fFeatures&FADF_UNKNOWN)
  {
    vt = VT_UNKNOWN;
  }
  else if (a.fFeatures&FADF_DISPATCH)
  {
    vt = VT_DISPATCH;
  }
  else if (a.fFeatures&FADF_VARIANT)
  {
    vt = VT_VARIANT;
  }
  else if (a.fFeatures&FADF_HAVEVARTYPE)
  {
    HRESULT hr = SafeArrayGetVartype(const_cast<SAFEARRAY*>(&a), &vt);
    if (FAILED(hr))
    {
      std::cerr << "[Unable to get type of array: " << errorFromCode(hr) << "]" << std::endl;
      std::cerr.flush();
      return Nan::Undefined();
    }
  }
  if (vt == VT_EMPTY)
  {
    std::cerr << "[Unable to get type of array (no useful flags set)]" << std::endl;
    std::cerr.flush();
    return Nan::Undefined();
  }
  const SAFEARRAYBOUND& bnds = a.rgsabound[0];
  if (a.cDims == 0)
  {
    return Nan::New<Array>(0);
  }
  else if (a.cDims == 1)
  {
    // fast array copy, using SafeArrayAccessData
    Local<Array> result = Nan::New<Array>(bnds.cElements);
    void* raw;
    HRESULT hr = SafeArrayAccessData(const_cast<SAFEARRAY*>(&a), &raw);
    if (FAILED(hr))
    {
      std::cerr << "[Unable to access array contents: " << errorFromCode(hr) << "]" << std::endl;
      std::cerr.flush();
      return Nan::Undefined();
    }
    for (unsigned idx = 0; idx < bnds.cElements; idx++)
    {
      Nan::Set(result, idx, ArrayPrimitiveToValue(thisObject, raw, vt, a.cbElements, idx));
    }
    hr = SafeArrayUnaccessData(const_cast<SAFEARRAY*>(&a));
    if (FAILED(hr))
    {
      std::cerr << "[Unable to release array contents: " << errorFromCode(hr) << "]" << std::endl;
      std::cerr.flush();
    }
    return result;
  } else {
    // slow array copy, using SafeArrayGetElement
    Local<Array> result = Nan::New<Array>(bnds.cElements);
    for (unsigned idx = 0; idx < bnds.cElements; idx++)
    {
      LONG srcIndex = idx + bnds.lLbound;
      Nan::Set(result, idx, ArrayToValueSlow(thisObject, a, vt, &srcIndex, 1));
    }
    return result;
  }
  OLETRACEOUT();
}

Local<Value> V8Variant::ArrayToValueSlow(Handle<Object> thisObject, const SAFEARRAY& a, VARTYPE vt, LONG* idices, unsigned numIdx)
{
  OLETRACEIN();
  OLETRACEFLUSH();
  if (numIdx <= a.cDims)
  {
    void* elm = alloca(a.cbElements);
    HRESULT hr = SafeArrayGetElement(const_cast<SAFEARRAY*>(&a), idices, elm);
    if (FAILED(hr))
    {
      std::cerr << "[Unable to access array element: " << errorFromCode(hr) << "]" << std::endl;
      std::cerr.flush();
      return Nan::Undefined();
    }
    return ArrayPrimitiveToValue(thisObject, elm, vt, a.cbElements, 0);
  } else {
    LONG* newIdx = (LONG*)alloca(sizeof(LONG) * (numIdx + 1));
    memcpy(newIdx, idices, sizeof(LONG)*numIdx);
    const SAFEARRAYBOUND& bnds = a.rgsabound[numIdx-1];
    Local<Array> result = Nan::New<Array>(bnds.cElements);
    for (unsigned idx = 0; idx < bnds.cElements; idx++)
    {
      newIdx[numIdx] = idx + bnds.lLbound;
      Nan::Set(result, idx, ArrayToValueSlow(thisObject, a, vt, newIdx, numIdx+1));
    }
    return result;
  }
  OLETRACEOUT();
}

Local<Value> V8Variant::VariantToValue(Handle<Object> thisObject, const VARIANT& v)
{
  OLETRACEIN();
  OLETRACEFLUSH();
  /*
  *  VT_CY               [V][T][P][S]  currency
  *  VT_UNKNOWN          [V][T]   [S]  IUnknown *
  *  VT_DECIMAL          [V][T]   [S]  16 byte fixed point
  *  VT_RECORD           [V]   [P][S]  user defined type
  */
  Local<Object> result;
  switch (v.vt)
  {
  case VT_EMPTY:
  case VT_NULL:
    return Nan::Null();
  case VT_ERROR:
    return Exception::Error(Nan::New<String>((const uint16_t*)errorFromCodeW(HRESULT_FROM_WIN32(v.scode)).c_str()).ToLocalChecked());
  case VT_ERROR | VT_BYREF:
    if (!v.ppdispVal) return Nan::Undefined(); // really shouldn't happen
    return Exception::Error(Nan::New<String>((const uint16_t*)errorFromCodeW(HRESULT_FROM_WIN32(*v.pscode)).c_str()).ToLocalChecked());
  case VT_DISPATCH:
    if (v.pdispVal == NULL) {
      return Nan::Null();
    } else {
      Handle<Object> vResult = V8Variant::CreateUndefined();
      V8Variant *o = V8Variant::Unwrap<V8Variant>(vResult);
      CHECK_V8V_UNDEFINED(o);
      VariantCopy(&o->ocv.v, &v); // copy rv value
      return vResult;
    }
  case VT_DISPATCH | VT_BYREF:
    if (!v.ppdispVal) return Nan::Undefined(); // really shouldn't happen
    if (*v.ppdispVal == NULL) {
      return Nan::Null();
    } else {
      Handle<Object> vResult = V8Variant::CreateUndefined();
      V8Variant *o = V8Variant::Unwrap<V8Variant>(vResult);
      CHECK_V8V_UNDEFINED(o);
      VariantCopy(&o->ocv.v, &v); // copy rv value
      VariantChangeType(&o->ocv.v, &o->ocv.v, 0, VT_DISPATCH);
      return vResult;
    }
  case VT_BOOL:
    return v.boolVal != VARIANT_FALSE ? Nan::True() : Nan::False();
  case VT_BOOL | VT_BYREF:
    if (!v.pboolVal) return Nan::Undefined(); // really shouldn't happen
    return *v.pboolVal != VARIANT_FALSE ? Nan::True() : Nan::False();
  case VT_I1:
    return Nan::New(v.cVal);
  case VT_I1 | VT_BYREF:
    if (!v.pcVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.pcVal);
  case VT_UI1:
    return Nan::New(v.bVal);
  case VT_UI1 | VT_BYREF:
    if (!v.pbVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.pbVal);
  case VT_I2:
    return Nan::New(v.iVal);
  case VT_I2 | VT_BYREF:
    if (!v.piVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.piVal);
  case VT_UI2:
    return Nan::New(v.uiVal);
  case VT_UI2 | VT_BYREF:
    if (!v.puiVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.puiVal);
  case VT_I4:
    return Nan::New(v.lVal);
  case VT_I4 | VT_BYREF:
    if (!v.plVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.plVal);
  case VT_UI4:
    return Nan::New((uint32_t)v.ulVal);
  case VT_UI4 | VT_BYREF:
    if (!v.pulVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New((uint32_t)*v.pulVal);
  case VT_INT:
    return Nan::New(v.intVal);
  case VT_INT | VT_BYREF:
    if (!v.pintVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.pintVal);
  case VT_UINT:
    return Nan::New(v.uintVal);
  case VT_UINT | VT_BYREF:
    if (!v.puintVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.puintVal);
  case VT_R4:
    return Nan::New(v.fltVal);
  case VT_R4 | VT_BYREF:
    if (!v.pfltVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.pfltVal);
  case VT_R8:
    return Nan::New(v.dblVal);
  case VT_R8 | VT_BYREF:
    if (!v.pdblVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New(*v.pdblVal);
  case VT_BSTR:
    if (!v.bstrVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New<String>((const uint16_t*)v.bstrVal).ToLocalChecked();
  case VT_BSTR | VT_BYREF:
    if (!v.pbstrVal || !*v.pbstrVal) return Nan::Undefined(); // really shouldn't happen
    return Nan::New<String>((const uint16_t*)*v.pbstrVal).ToLocalChecked();
  case VT_DATE:
    return OLEDateToObject(v.date);
  case VT_DATE | VT_BYREF:
    if (!v.pdate) return Nan::Undefined(); // really shouldn't happen
    return OLEDateToObject(*v.pdate);
  case VT_VARIANT | VT_BYREF:
    if (!v.pvarVal) return Nan::Undefined(); // really shouldn't happen
    return VariantToValue(thisObject, *v.pvarVal);
  case VT_SAFEARRAY:
    if (!v.parray) return Nan::Undefined(); // really shouldn't happen
    return ArrayToValue(thisObject, *v.parray);
  case VT_SAFEARRAY | VT_BYREF:
    if (!v.pparray || !*v.pparray) return Nan::Undefined(); // really shouldn't happen
    return ArrayToValue(thisObject, **v.pparray);
  default:
    if (v.vt & VT_ARRAY)
    {
      if (v.vt & VT_BYREF)
      {
        if (!v.pparray || !*v.pparray) return Nan::Undefined(); // really shouldn't happen
        return ArrayToValue(thisObject, **v.pparray);
      } else {
        if (!v.parray) return Nan::Undefined(); // really shouldn't happen
        return ArrayToValue(thisObject, *v.parray);
      }
    } else {
      Handle<Value> s = INSTANCE_CALL(thisObject, "vtName", 0, NULL);
      std::cerr << "[unknown type " << v.vt << ":" << *String::Utf8Value(s);
      std::cerr << " (not implemented now)]" << std::endl;
      std::cerr.flush();
    }
    return Nan::Undefined();
  }
  OLETRACEOUT();
}

static std::string GetName(ITypeInfo *typeinfo, MEMBERID id) {
  BSTR name;
  UINT numNames = 0;
  typeinfo->GetNames(id, &name, 1, &numNames);
  if (numNames > 0) {
    return BSTR2MBCS(name);
  }
  return "";
}

/**
 * Like OLEValue, but goes one step further and reduces IDispatch objects
 * to actual JS objects. This enables things like console.log().
 **/
NAN_METHOD(V8Variant::OLEPrimitiveValue) {
  Local<Object> thisObject = info.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  OLETRACEVT(thisObject);
  OLETRACEFLUSH();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  VARIANT& v = v8v->ocv.v;
  IDispatch* dispatch = NULL;
  switch (v.vt)
  {
  case VT_DISPATCH:
    dispatch = v.pdispVal;
    break;
  case VT_DISPATCH | VT_BYREF:
    if (!v.ppdispVal) return info.GetReturnValue().SetUndefined(); // really shouldn't happen
    dispatch = *v.ppdispVal;
    break;
  default:
    return info.GetReturnValue().Set(VariantToValue(info.This(), v));
  }
  if (dispatch == NULL) {
    return info.GetReturnValue().SetNull();
  }
  Local<Object> object = Nan::New<Object>();
  ITypeInfo *typeinfo = NULL;
  HRESULT hr = dispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &typeinfo);
  if (typeinfo) {
    TYPEATTR* typeattr;
    BASSERT(SUCCEEDED(typeinfo->GetTypeAttr(&typeattr)));
    for (int i = 0; i < typeattr->cFuncs; i++) {
      FUNCDESC *funcdesc;
      typeinfo->GetFuncDesc(i, &funcdesc);
      if (funcdesc->invkind != INVOKE_FUNC) {
        Nan::Set(object, Nan::New(GetName(typeinfo, funcdesc->memid)).ToLocalChecked(), Nan::New("Function").ToLocalChecked());
      }
      typeinfo->ReleaseFuncDesc(funcdesc);
    }
    for (int i = 0; i < typeattr->cVars; i++) {
      VARDESC *vardesc;
      typeinfo->GetVarDesc(i, &vardesc);
      Nan::Set(object, Nan::New(GetName(typeinfo, vardesc->memid)).ToLocalChecked(), Nan::New("Variable").ToLocalChecked());
      typeinfo->ReleaseVarDesc(vardesc);
    }
    typeinfo->ReleaseTypeAttr(typeattr);
  }
  return info.GetReturnValue().Set(object);
}

Handle<Object> V8Variant::CreateUndefined(void)
{
  DISPFUNCIN();
  Local<FunctionTemplate> localClazz = Nan::New(clazz);
  Local<Object> instance = Nan::NewInstance(Nan::GetFunction(localClazz).ToLocalChecked(), 0, NULL).ToLocalChecked();
  DISPFUNCOUT();
  return instance;
}

NAN_METHOD(V8Variant::New)
{
  DISPFUNCIN();
  if(!info.IsConstructCall())
    return Nan::ThrowTypeError("Use the new operator to create new V8Variant objects");
  Local<Object> thisObject = info.This();
  V8Variant *v = new V8Variant(); // must catch exception
  CHECK_V8V(v);
  v->Wrap(thisObject); // InternalField[0]
  DISPFUNCOUT();
  return info.GetReturnValue().Set(info.This());
}

Handle<Value> V8Variant::OLEFlushCarryOver(Handle<Value> v)
{
  OLETRACEIN();
  OLETRACEVT_UNDEFINED(Nan::To<Object>(v).ToLocalChecked());
  Handle<Value> result = Nan::Undefined();
  V8Variant *v8v = node::ObjectWrap::Unwrap<V8Variant>(Nan::To<Object>(v).ToLocalChecked());
  if(v8v->property_carryover.empty()){
    std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
    std::cerr.flush();
    // *** or throw exception
  }else{
    const char *name = v8v->property_carryover.c_str();
    {
      OLETRACEPREARGV(Nan::New(name).ToLocalChecked());
      OLETRACEARGV();
    }
    OLETRACEFLUSH();
    Handle<Value> argv[] = {Nan::New(name).ToLocalChecked(), Nan::New<Array>(0)};
    int argc = sizeof(argv) / sizeof(argv[0]); // == 2
    v8v->property_carryover.clear();
    result = INSTANCE_CALL(Nan::To<Object>(v).ToLocalChecked(), "call", argc, argv);
    OCVariant *rv = V8Variant::CreateOCVariant(result);
    CHECK_OCV_UNDEFINED(rv);
    V8Variant *o = V8Variant::Unwrap<V8Variant>(Nan::To<Object>(v).ToLocalChecked());
    CHECK_V8V_UNDEFINED(o);
    o->ocv = *rv; // copy rv value
    delete rv;
  }
  OLETRACEOUT();
  return result;
}

template<bool isCall>
NAN_METHOD(OLEInvoke)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS(info, 1, av0, av1);
  String::Utf8Value u8s(av0);
  BSTR bName = MBCS2BSTR(*u8s);
  if (!bName) return Nan::ThrowError(NewOleException(E_OUTOFMEMORY));
  Array *a = Array::Cast(*av1);
  uint32_t argLen = a->Length();
  OCVariant **argchain = argLen ? (OCVariant**)alloca(sizeof(OCVariant*)*argLen) : NULL;
  for(uint32_t i = 0; i < argLen; ++i){
    OCVariant *o = V8Variant::CreateOCVariant(ARRAY_AT(a, i));
    CHECK_OCV(o);
    argchain[i] = o;
  }
  ErrorInfo errInfo;
  OCVariant rv;
  HRESULT hr = isCall ? // argchain will be deleted automatically
    v8v->ocv.invoke(bName, &rv, errInfo, argchain, argLen) : v8v->ocv.getProp(bName, rv, errInfo, argchain, argLen);
  ::SysFreeString(bName);
  if (FAILED(hr))
  {
    return Nan::ThrowError(NewOleException(hr, errInfo));
  }
//  Handle<Value> result = INSTANCE_CALL(vResult, "toValue", 0, NULL);
  Local<Value> vResult = V8Variant::VariantToValue(info.This(), rv.v);
  OLETRACEOUT();
  if (!vResult->IsUndefined())
  {
    return info.GetReturnValue().Set(vResult);
  }
}

NAN_METHOD(V8Variant::OLECall)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OLEInvoke<true>(info);
  OLETRACEOUT();
}

NAN_METHOD(V8Variant::OLEGet)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  OLEInvoke<false>(info);
  OLETRACEOUT();
}

NAN_METHOD(V8Variant::OLESet)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  OLETRACEARGS();
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  OLE_PROCESS_CARRY_OVER(thisObject);
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(info.This());
  CHECK_V8V(v8v);
  Handle<Value> av0, av1;
  CHECK_OLE_ARGS(info, 2, av0, av1);
  OCVariant *argchain = V8Variant::CreateOCVariant(av1);
  if(!argchain)
    return Nan::ThrowTypeError(__FUNCTION__ " the second argument is not valid (null OCVariant)");
  String::Utf8Value u8s(av0);
  BSTR bName = MBCS2BSTR(*u8s);
  if (!bName) return Nan::ThrowError(NewOleException(E_OUTOFMEMORY));
  ErrorInfo errInfo;
  HRESULT hr = v8v->ocv.putProp(bName, errInfo, &argchain, 1); // argchain will be deleted automatically
  ::SysFreeString(bName);
  if (FAILED(hr))
  {
    return Nan::ThrowError(NewOleException(hr, errInfo));
  }
  OLETRACEOUT();
}

NAN_METHOD(V8Variant::OLECallComplete)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  Handle<Value> result = Nan::Undefined();
  V8Variant *v8v = node::ObjectWrap::Unwrap<V8Variant>(info.This());
  if(v8v->property_carryover.empty()){
    std::cerr << " *** carryover empty *** " << __FUNCTION__ << std::endl;
    std::cerr.flush();
    // *** or throw exception
  }else{
    const char *name = v8v->property_carryover.c_str();
    {
      OLETRACEPREARGV(Nan::New(name).ToLocalChecked());
      OLETRACEARGV();
    }
    OLETRACEARGS();
    OLETRACEFLUSH();
    int argLen = info.Length();
    Handle<Array> a = Nan::New<Array>(argLen);
    for(int i = 0; i < argLen; ++i) ARRAY_SET(a, i, info[i]);
    Handle<Value> argv[] = {Nan::New(name).ToLocalChecked(), a};
    int argc = sizeof(argv) / sizeof(argv[0]); // == 2
    v8v->property_carryover.clear();
    result = INSTANCE_CALL(info.This(), "call", argc, argv);
  }
//_
//Handle<Value> r = INSTANCE_CALL(Handle<Object>::Cast(v), "toValue", 0, NULL);
  OLETRACEOUT();
  return info.GetReturnValue().Set(result);
}

NAN_PROPERTY_GETTER(V8Variant::OLEGetAttr)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  {
    OLETRACEPREARGV(property);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  String::Utf8Value u8name(property);
  Local<Object> thisObject = info.This();

  // Why GetAttr comes twice for () in the third loop instead of CallComplete ?
  // Because of the Crankshaft v8's run-time optimizer ?
  {
    V8Variant *v8v = node::ObjectWrap::Unwrap<V8Variant>(thisObject);
    if(!v8v->property_carryover.empty()){
      if(v8v->property_carryover == *u8name){
        OLETRACEOUT();
        return info.GetReturnValue().Set(thisObject); // through it
      }
    }
  }

#if(0)
  if(std::string("call") == *u8name || std::string("get") == *u8name
  || std::string("_") == *u8name || std::string("toValue") == *u8name
//|| std::string("valueOf") == *u8name || std::string("toString") == *u8name
  ){
    OLE_PROCESS_CARRY_OVER(thisObject);
  }
#else
  if(std::string("set") != *u8name
  && std::string("toBoolean") != *u8name
  && std::string("toInt32") != *u8name && std::string("toInt64") != *u8name
  && std::string("toNumber") != *u8name && std::string("toDate") != *u8name
  && std::string("toUtf8") != *u8name
  && std::string("inspect") != *u8name && std::string("constructor") != *u8name
  && std::string("valueOf") != *u8name && std::string("toString") != *u8name
  && std::string("toLocaleString") != *u8name
  && std::string("toJSON") != *u8name
  && std::string("hasOwnProperty") != *u8name
  && std::string("isPrototypeOf") != *u8name
  && std::string("propertyIsEnumerable") != *u8name
//&& std::string("_") != *u8name
  ){
    OLE_PROCESS_CARRY_OVER(thisObject);
  }
#endif
  OLETRACEVT(thisObject);
  // Can't use INSTANCE_CALL here. (recursion itself)
  // So it returns Object's fundamental function and custom function:
  //   inspect ?, constructor valueOf toString toLocaleString
  //   hasOwnProperty isPrototypeOf propertyIsEnumerable
  static fundamental_attr fundamentals[] = {
    {0, "call", OLECall}, {0, "get", OLEGet}, {0, "set", OLESet},
    {0, "isA", OLEIsA}, {0, "vtName", OLEVTName}, // {"vt_names", ???},
    {!0, "toBoolean", OLEBoolean},
    {!0, "toInt32", OLEInt32},
    {!0, "toNumber", OLENumber}, {!0, "toDate", OLEDate},
    {!0, "toUtf8", OLEUtf8},
    {0, "toValue", OLEValue},
    {0, "inspect", OLEPrimitiveValue}, {0, "constructor", NULL}, {0, "valueOf", OLEPrimitiveValue},
    {0, "toString", OLEPrimitiveValue}, {0, "toLocaleString", OLEPrimitiveValue},
    {0, "toJSON", OLEPrimitiveValue},
    {0, "hasOwnProperty", NULL}, {0, "isPrototypeOf", NULL},
    {0, "propertyIsEnumerable", NULL}
  };
  for(int i = 0; i < sizeof(fundamentals) / sizeof(fundamentals[0]); ++i){
    if(std::string(fundamentals[i].name) != *u8name) continue;
    if(fundamentals[i].obsoleted){
      std::cerr << " *** ## [." << fundamentals[i].name;
      std::cerr << "()] is obsoleted. ## ***" << std::endl;
      std::cerr.flush();
    }
    OLETRACEFLUSH();
    OLETRACEOUT();
    return info.GetReturnValue().Set(Nan::GetFunction(Nan::New<FunctionTemplate>(
      fundamentals[i].func, thisObject)).ToLocalChecked());
  }
  if(std::string("_") == *u8name){ // through it when "_"
#if(0)
    std::cerr << " *** ## [._] is obsoleted. ## ***" << std::endl;
    std::cerr.flush();
#endif
  }else{
    Handle<Object> vResult = V8Variant::CreateUndefined(); // uses much memory
    OCVariant *rv = V8Variant::CreateOCVariant(thisObject);
    CHECK_OCV(rv);
    V8Variant *o = V8Variant::Unwrap<V8Variant>(vResult);
    CHECK_V8V(o);
    o->ocv = *rv; // copy rv value
    delete rv;
    V8Variant *v8v = node::ObjectWrap::Unwrap<V8Variant>(vResult);
    v8v->property_carryover.assign(*u8name);
    OLETRACEPREARGV(property);
    OLETRACEARGV();
    OLETRACEFLUSH();
    OLETRACEOUT();
    return info.GetReturnValue().Set(vResult); // convert and hold it (uses much memory)
  }
  OLETRACEFLUSH();
  OLETRACEOUT();
  return info.GetReturnValue().Set(thisObject); // through it
}

NAN_PROPERTY_SETTER(V8Variant::OLESetAttr)
{
  OLETRACEIN();
  OLETRACEVT(info.This());
  Handle<Value> argv[] = {property, value};
  int argc = sizeof(argv) / sizeof(argv[0]);
  OLETRACEARGV();
  OLETRACEFLUSH();
  Handle<Value> r = INSTANCE_CALL(info.This(), "set", argc, argv);
  OLETRACEOUT();
  return info.GetReturnValue().Set(r);
}

NAN_METHOD(V8Variant::Finalize)
{
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
  std::cerr.flush();
#endif
  V8Variant *v = node::ObjectWrap::Unwrap<V8Variant>(info.This());
  if(v) v->Finalize();
  DISPFUNCOUT();
  return info.GetReturnValue().Set(info.This());
}

void V8Variant::Finalize()
{
  if(!finalized)
  {
    ocv.Clear();
    finalized = true;
  }
}

} // namespace node_win32ole
