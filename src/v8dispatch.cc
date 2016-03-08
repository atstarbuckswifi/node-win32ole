/*
  v8dispatch.cc
*/

#include "v8dispatch.h"
#include <node.h>
#include <nan.h>
#include "v8dispmember.h"
#include "v8variant.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Nan::Persistent<FunctionTemplate> V8Dispatch::clazz;

void V8Dispatch::Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target)
{
  Nan::HandleScope scope;
  Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(1);
  t->SetClassName(Nan::New("V8Dispatch").ToLocalChecked());
  Nan::SetPrototypeMethod(t, "valueOf", OLEPrimitiveValue);
  Nan::SetPrototypeMethod(t, "toString", OLEStringValue);
  Nan::SetPrototypeMethod(t, "toLocaleString", OLELocaleStringValue);
//  Nan::SetPrototypeMethod(t, "New", New);
/*
 In ParseUnaryExpression() < v8/src/parser.cc >
 v8::Object::ToBoolean() is called directly for unary operator '!'
 instead of v8::Object::valueOf()
 so NamedPropertyHandler will not be called
 Local<Boolean> ToBoolean(); // How to fake ? override v8::Value::ToBoolean
*/
  Local<ObjectTemplate> instancetpl = t->InstanceTemplate();
  //Nan::SetCallAsFunctionHandler(instancetpl, OLECallComplete);
  Nan::SetNamedPropertyHandler(instancetpl, OLEGetAttr, OLESetAttr);
  Nan::SetIndexedPropertyHandler(instancetpl, OLEGetIdxAttr, OLESetIdxAttr);
  Nan::SetPrototypeMethod(t, "Finalize", Finalize);
  Nan::Set(target, Nan::New("V8Dispatch").ToLocalChecked(), t->GetFunction());
  clazz.Reset(t);
}

NAN_METHOD(V8Dispatch::OLEValue)
{
  OLETRACEIN();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(info.This());
  CHECK_V8(V8Dispatch, vThis);
  Local<Value> vResult = vThis->OLEGet(DISPID_VALUE);
  if (!vResult->IsUndefined())
  {
    return info.GetReturnValue().Set(vResult);
  }
  OLETRACEOUT();
}

Local<Value> V8Dispatch::resolveValueChain(Local<Object> thisObject, const char* prop)
{
  OLETRACEIN();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(thisObject);
  CHECK_V8_UNDEFINED(V8Dispatch, vThis);
  Local<Value> vResult = vThis->OLEGet(DISPID_VALUE);
  if (vResult->IsUndefined()) return Nan::Undefined();
  vResult = INSTANCE_CALL(Nan::To<Object>(vResult).ToLocalChecked(), prop, 0, NULL);
  return vResult;
}

/**
* Like OLEValue, but goes one step further and reduces IDispatch objects
* to actual JS objects. This enables things like console.log().
**/
NAN_METHOD(V8Dispatch::OLEPrimitiveValue)
{
  OLETRACEIN();
  Local<Value> vResult = resolveValueChain(info.This(), "valueOf");
  if (vResult->IsUndefined()) return;
  OLETRACEOUT();
  return info.GetReturnValue().Set(vResult);
}

NAN_METHOD(V8Dispatch::OLEStringValue)
{
  OLETRACEIN();
  Local<Value> vResult = resolveValueChain(info.This(), "toString");
  if (vResult->IsUndefined()) return;
  OLETRACEOUT();
  return info.GetReturnValue().Set(vResult);
}

NAN_METHOD(V8Dispatch::OLELocaleStringValue)
{
  OLETRACEIN();
  Local<Value> vResult = resolveValueChain(info.This(), "toLocaleString");
  if (vResult->IsUndefined()) return;
  OLETRACEOUT();
  return info.GetReturnValue().Set(vResult);
}

Handle<Object> V8Dispatch::CreateNew(IDispatch* disp)
{
  DISPFUNCIN();
  Local<FunctionTemplate> localClazz = Nan::New(clazz);
  Local<Object> instance = Nan::NewInstance(Nan::GetFunction(localClazz).ToLocalChecked(), 0, NULL).ToLocalChecked();
  if (disp)
  {
    V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(instance);
    vThis->ocd.disp = disp;
    vThis->ocd.disp->AddRef();
  }
  DISPFUNCOUT();
  return instance;
}

NAN_METHOD(V8Dispatch::New)
{
  DISPFUNCIN();
  if(!info.IsConstructCall())
    return Nan::ThrowTypeError("Use the new operator to create new V8Dispatch objects");
  Local<Object> thisObject = info.This();
  V8Dispatch *v = new V8Dispatch(); // must catch exception
  CHECK_V8(V8Dispatch, v);
  v->Wrap(thisObject); // InternalField[0]
  DISPFUNCOUT();
  return info.GetReturnValue().Set(info.This());
}

Local<Value> V8Dispatch::OLECall(DISPID propID, Nan::NAN_METHOD_ARGS_TYPE info)
{
  DISPFUNCIN();
  int argc = info.Length();
  Local<Value>* argv = (Local<Value>*)alloca(sizeof(Local<Value>) * argc);
  for (int idx = 0; idx < argc; ++idx)
  {
    *(new(argv + idx) Local<Value>) = info[idx];
  }
  Local<Value> hResult = OLECall(propID, argc, argv);
  for (int idx = 0; idx < argc; ++idx)
  {
    (argv + idx)->~Local<Value>();
  }
  DISPFUNCOUT();
  return hResult;
}

Local<Value> V8Dispatch::OLECall(DISPID propID, int argc, Local<Value> argv[])
{
  OLETRACEIN();
  OCVariant **argchain = argc ? (OCVariant**)alloca(sizeof(OCVariant*) * argc) : NULL;
  for (int i = 0; i < argc; ++i) {
    OCVariant *o = V8Variant::ValueToVariant(argv[i]);
    if (!o) return Nan::Undefined();
    argchain[i] = o;
  }
  ErrorInfo errInfo;
  OCVariant rv;
  HRESULT hr = ocd.invoke(propID, &rv, errInfo, argchain, argc); // argchain will be deleted automatically
  if (FAILED(hr))
  {
    Nan::ThrowError(NewOleException(hr, errInfo));
    return Nan::Undefined();
  }
  Local<Value> vResult = V8Variant::VariantToValue(rv.v);
  OLETRACEOUT();
  return vResult;
}

Local<Value> V8Dispatch::OLEGet(DISPID propID, int argc, Local<Value> argv[])
{
  OLETRACEIN();
  OCVariant **argchain = argc ? (OCVariant**)alloca(sizeof(OCVariant*) * argc) : NULL;
  for (int i = 0; i < argc; ++i) {
    OCVariant *o = V8Variant::ValueToVariant(argv[i]);
    if (!o) return Nan::Undefined();
    argchain[i] = o;
  }
  ErrorInfo errInfo;
  OCVariant rv;
  HRESULT hr = ocd.getProp(propID, rv, errInfo, argchain, argc); // argchain will be deleted automatically
  if (FAILED(hr))
  {
    Nan::ThrowError(NewOleException(hr, errInfo));
    return Nan::Undefined();
  }
  Local<Value> vResult = V8Variant::VariantToValue(rv.v);
  OLETRACEOUT();
  return vResult;
}

bool V8Dispatch::OLESet(DISPID propID, int argc, Local<Value> argv[])
{
  OLETRACEIN();
  OCVariant **argchain = argc ? (OCVariant**)alloca(sizeof(OCVariant*) * argc) : NULL;
  for (int i = 0; i < argc; ++i) {
    OCVariant *o = V8Variant::ValueToVariant(argv[i]);
    if (!o) return false;
    argchain[i] = o;
  }
  ErrorInfo errInfo;
  HRESULT hr = ocd.putProp(propID, errInfo, argchain, argc); // argchain will be deleted automatically
  if (FAILED(hr))
  {
    Nan::ThrowError(NewOleException(hr, errInfo));
    return false;
  }
  OLETRACEOUT();
  return true;
}

NAN_PROPERTY_GETTER(V8Dispatch::OLEGetAttr)
{
  OLETRACEIN();
  {
    OLETRACEPREARGV(property);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(info.This());
  CHECK_V8(V8Dispatch, vThis);

  String::Value vProperty(property);
  if (wcscmp((const wchar_t*)*vProperty, L"_") == 0)
  {
    Local<Value> vResult = vThis->OLEGet(DISPID_VALUE);
    if (vResult->IsUndefined()) return; // exception?
    return info.GetReturnValue().Set(vResult);
  }

  // try to resolve this as a member of our object
  String::Value vName(property);
  BSTR bName = ::SysAllocString((const wchar_t*)*vName);
  if (bName)
  {
    DISPID dispID;
    HRESULT hr = vThis->ocd.disp->GetIDsOfNames(IID_NULL, &bName, 1, LOCALE_USER_DEFAULT, &dispID);
    ::SysFreeString(bName);
    if (SUCCEEDED(hr))
    {
      Handle<Object> vDispMember = V8DispMember::CreateNew(thisObject, dispID);
      return info.GetReturnValue().Set(vDispMember);
    }
  }

  // try to retrieve it as an existing property of the js object
  Nan::MaybeLocal<Value> mLocal = Nan::GetRealNamedProperty(thisObject, property);
  if (!mLocal.IsEmpty())
  {
    Local<Value> hLocal = mLocal.ToLocalChecked();
    if (!hLocal->IsUndefined() && !hLocal->IsNull())
    { // this is a javascript property
      return info.GetReturnValue().Set(hLocal);
    }
  }

  OLETRACEOUT();
}

NAN_PROPERTY_SETTER(V8Dispatch::OLESetAttr)
{
  OLETRACEIN();
  {
    Handle<Value> argv[] = { property, value };
    int argc = sizeof(argv) / sizeof(argv[0]);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(info.This());
  CHECK_V8(V8Dispatch, vThis);

  // try to resolve this as a member of our object
  String::Value vName(property);
  BSTR bName = ::SysAllocString((const wchar_t*)*vName);
  if (bName)
  {
    DISPID dispID;
    HRESULT hr = vThis->ocd.disp->GetIDsOfNames(IID_NULL, &bName, 1, LOCALE_USER_DEFAULT, &dispID);
    ::SysFreeString(bName);
    if (SUCCEEDED(hr))
    {
      bool bResult = vThis->OLESet(dispID, 1, &value);
      return info.GetReturnValue().Set(bResult);
    }
  }

  // if it is an existing property of the js object, permit it to be set
  Nan::MaybeLocal<Value> mLocal = Nan::GetRealNamedProperty(thisObject, property);
  if (!mLocal.IsEmpty())
  {
    Local<Value> hLocal = mLocal.ToLocalChecked();
    if (!hLocal->IsUndefined() && !hLocal->IsNull())
    { // this is a javascript property
      Nan::Maybe<bool> bResult = Nan::ForceSet(thisObject, property, value);
      if (!bResult.IsNothing()) return info.GetReturnValue().Set(bResult.IsJust());
      return;
    }
  }

  // if it's not a known COM or js property, then we don't want the user creating arbitrary expandos
  OLETRACEOUT();
  return Nan::ThrowError("Cannot assign new properties to this object");
}

NAN_INDEX_GETTER(V8Dispatch::OLEGetIdxAttr)
{
  OLETRACEIN();
  {
    OLETRACEPREARGV(Nan::New(index));
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(info.This());
  CHECK_V8(V8Dispatch, vThis);

  // attempt to resolve ourselves by assuming the default property is indexed
  Handle<Value> argv[] = { Nan::New(index) };
  int argc = sizeof(argv) / sizeof(argv[0]); // == 1
  Local<Value> vResult = vThis->OLEGet(DISPID_VALUE, argc, argv);
  if (!vResult->IsUndefined())
  {
    return info.GetReturnValue().Set(vResult);
  }
  OLETRACEOUT();
}

NAN_INDEX_SETTER(V8Dispatch::OLESetIdxAttr)
{
  OLETRACEIN();
  {
    Handle<Value> argv[] = { Nan::New(index), value };
    int argc = sizeof(argv) / sizeof(argv[0]);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(info.This());
  CHECK_V8(V8Dispatch, vThis);

  // attempt to resolve ourselves by assuming the default property is indexed
  Handle<Value> argv[] = { Nan::New(index), value };
  int argc = sizeof(argv) / sizeof(argv[0]); // == 1
  bool bResult = vThis->OLESet(DISPID_VALUE, argc, argv);
  if (bResult)
  {
    return info.GetReturnValue().Set(true);
  }
  OLETRACEOUT();
}

// NAN_PROPERTY_ENUMERATOR
// NAN_PROPERTY_DELETER
// NAN_PROPERTY_QUERY
// NAN_INDEX_ENUMERATOR
// NAN_INDEX_DELETER
// NAN_INDEX_QUERY

NAN_METHOD(V8Dispatch::Finalize)
{
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
  std::cerr.flush();
#endif
  V8Dispatch *v = V8Dispatch::Unwrap<V8Dispatch>(info.This());
  if(v) v->Finalize();
  DISPFUNCOUT();
  return info.GetReturnValue().Set(info.This());
}

void V8Dispatch::Finalize()
{
  if(!finalized)
  {
    ocd.Clear();
    finalized = true;
  }
}

} // namespace node_win32ole
