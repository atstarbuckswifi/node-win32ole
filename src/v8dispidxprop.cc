/*
  v8dispidxprop.cc
*/

#include "v8dispidxprop.h"
#include <node.h>
#include <nan.h>
#include "v8dispatch.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Nan::Persistent<FunctionTemplate> V8DispIdxProperty::clazz;

void V8DispIdxProperty::Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target)
{
  Nan::HandleScope scope;
  Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(1);
  t->SetClassName(Nan::New("V8DispIdxProperty").ToLocalChecked());
  Nan::SetPrototypeMethod(t, "valueOf", OLEStringValue);
  Nan::SetPrototypeMethod(t, "toString", OLEStringValue);
  Nan::SetPrototypeMethod(t, "toLocaleString", OLEStringValue);

  Local<ObjectTemplate> instancetpl = t->InstanceTemplate();
  Nan::SetNamedPropertyHandler(instancetpl, OLEGetAttr, OLESetAttr);
  Nan::SetIndexedPropertyHandler(instancetpl, OLEGetIdxAttr, OLESetIdxAttr);
  Nan::Set(target, Nan::New("V8DispIdxProperty").ToLocalChecked(), t->GetFunction());
  clazz.Reset(t);
}

NAN_METHOD(V8DispIdxProperty::OLEStringValue)
{
  OLETRACEIN();
  Local<Object> thisObject = info.This();
  V8DispIdxProperty *vThis = V8DispIdxProperty::Unwrap<V8DispIdxProperty>(thisObject);
  CHECK_V8(V8DispIdxProperty, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);
  std::wstring fullName = (vDisp->m_typeName.empty() ? L"Object" : vDisp->m_typeName) + L"." + vThis->name;
  OLETRACEOUT();
  return info.GetReturnValue().Set(Nan::New((const uint16_t*)fullName.c_str()).ToLocalChecked());
}

MaybeLocal<Object> V8DispIdxProperty::CreateNew(Handle<Object> dispatch, Handle<String> property, DISPID id)
{
  DISPFUNCIN();
  Local<FunctionTemplate> localClazz = Nan::New(clazz);
  Local<v8::Value> args[] = { dispatch, property, Nan::New<Int32>(id) };
  int argc = sizeof(args) / sizeof(args[0]); // == 3
  return Nan::NewInstance(Nan::GetFunction(localClazz).ToLocalChecked(), argc, args);
  DISPFUNCOUT();
}

NAN_METHOD(V8DispIdxProperty::New)
{
  DISPFUNCIN();
  if(!info.IsConstructCall())
    return Nan::ThrowTypeError("Use the new operator to create new V8DispMember objects");
  if (info.Length() != 3)
    return Nan::ThrowTypeError("Must be constructed with (V8Dispatch, Name, DISPID)");
  Local<FunctionTemplate> v8DispatchClazz = Nan::New(V8Dispatch::clazz);
  if (!info[0]->IsObject() || !v8DispatchClazz->HasInstance(info[0]))
    return Nan::ThrowTypeError("First argument must be a V8Dispatch object");
  if(!info[1]->IsString() && !info[2]->IsStringObject())
    return Nan::ThrowTypeError("Second argument must be a string");
  Handle<String> hName = Nan::To<String>(info[1]).ToLocalChecked();
  if(!info[2]->IsInt32())
    return Nan::ThrowTypeError("Third argument must be a DISPID");
  Nan::MaybeLocal<Int32> maybeDispId = Nan::To<Int32>(info[2]);
  if(maybeDispId.IsEmpty())
    return Nan::ThrowTypeError("Second argument must be a DISPID");

  String::Value vName(hName);
  DISPID id = maybeDispId.ToLocalChecked()->Int32Value();

  V8DispIdxProperty *v = new V8DispIdxProperty((const wchar_t*)*vName, id); // must catch exception
  CHECK_V8(V8DispMember, v);
  Local<Object> thisObject = info.This();
  v->Wrap(thisObject); // InternalField[0]
  Nan::ForceSet(thisObject, Nan::New<String>("_dispatch").ToLocalChecked(), info[0]);
  DISPFUNCOUT();
  return info.GetReturnValue().Set(thisObject);
}

V8Dispatch* V8DispIdxProperty::getDispatch(Local<Object> dispMemberObj)
{
  Nan::MaybeLocal<Value> mpropDisp = Nan::GetRealNamedProperty(dispMemberObj, Nan::New<String>("_dispatch").ToLocalChecked());
  if (mpropDisp.IsEmpty())
    return NULL;
  Local<Value> propDisp = mpropDisp.ToLocalChecked();
  Local<FunctionTemplate> v8DispatchClazz = Nan::New(V8Dispatch::clazz);
  if (!propDisp->IsObject() || !v8DispatchClazz->HasInstance(propDisp))
    return NULL;
  return V8Dispatch::Unwrap<V8Dispatch>(Local<Object>::Cast(propDisp));
}

NAN_PROPERTY_GETTER(V8DispIdxProperty::OLEGetAttr)
{
  OLETRACEIN();
  {
    OLETRACEPREARGV(property);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();

  // first let's try to retrieve it as an existing property of the js object.  This is documented to bypass ourselves
  Nan::MaybeLocal<Value> mLocal = Nan::GetRealNamedProperty(thisObject, property);
  if (!mLocal.IsEmpty())
  {
    Local<Value> hLocal = mLocal.ToLocalChecked();
    if (!hLocal->IsUndefined() && !hLocal->IsNull())
    { // this is a javascript property
      return info.GetReturnValue().Set(hLocal);
    }
  }

  // Resolve ourselves as a property reference
  V8DispIdxProperty *vThis = V8DispIdxProperty::Unwrap<V8DispIdxProperty>(thisObject);
  CHECK_V8(V8DispIdxProperty, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);

  Handle<Value> argv[] = { property };
  int argc = sizeof(argv) / sizeof(argv[0]); // == 1
  Local<Value> vResult = vDisp->OLEGet(vThis->memberId, argc, argv);
  if (!vResult->IsUndefined())
  {
    return info.GetReturnValue().Set(vResult);
  }

  OLETRACEOUT();
}

NAN_PROPERTY_SETTER(V8DispIdxProperty::OLESetAttr)
{
  OLETRACEIN();
  {
    Handle<Value> argv[] = { property, value };
    int argc = sizeof(argv) / sizeof(argv[0]);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();

  // first check to see if it is an existing property of the js object.  This is documented to bypass ourselves
  Nan::MaybeLocal<Value> mLocal = Nan::GetRealNamedProperty(thisObject, property);
  if (!mLocal.IsEmpty())
  {
    Local<Value> hLocal = mLocal.ToLocalChecked();
    if (!hLocal->IsUndefined() && !hLocal->IsNull())
    { // this is a javascript property
      Nan::Maybe<bool> bResult = Nan::ForceSet(thisObject, property, value);
      if(!bResult.IsNothing()) return info.GetReturnValue().Set(bResult.IsJust());
      return;
    }
  }

  // Resolve ourselves as a property reference
  V8DispIdxProperty *vThis = V8DispIdxProperty::Unwrap<V8DispIdxProperty>(thisObject);
  CHECK_V8(V8DispIdxProperty, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);

  Handle<Value> argv[] = { property, value };
  int argc = sizeof(argv) / sizeof(argv[0]); // == 2
  bool bResult = vDisp->OLESet(vThis->memberId, argc, argv);
  if (bResult) return info.GetReturnValue().Set(true);

  OLETRACEOUT();
}

NAN_INDEX_GETTER(V8DispIdxProperty::OLEGetIdxAttr)
{
  OLETRACEIN();
  {
    OLETRACEPREARGV(Nan::New(index));
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();

  // Resolve ourselves as an indexed property reference
  V8DispIdxProperty *vThis = V8DispIdxProperty::Unwrap<V8DispIdxProperty>(thisObject);
  CHECK_V8(V8DispMember, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);

  Handle<Value> argv[] = { Nan::New(index) };
  int argc = sizeof(argv) / sizeof(argv[0]); // == 1
  Local<Value> vResult = vDisp->OLEGet(vThis->memberId, argc, argv);
  if (!vResult->IsUndefined())
  {
    return info.GetReturnValue().Set(vResult);
  }
  OLETRACEOUT();
}

NAN_INDEX_SETTER(V8DispIdxProperty::OLESetIdxAttr)
{
  OLETRACEIN();
  {
    Handle<Value> argv[] = { Nan::New(index), value };
    int argc = sizeof(argv) / sizeof(argv[0]);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();

  // Resolve ourselves as an indexed property reference
  V8DispIdxProperty *vThis = V8DispIdxProperty::Unwrap<V8DispIdxProperty>(thisObject);
  CHECK_V8(V8DispMember, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);

  Handle<Value> argv[] = { Nan::New(index), value };
  int argc = sizeof(argv) / sizeof(argv[0]); // == 2
  bool bResult = vDisp->OLESet(vThis->memberId, argc, argv);
  if (bResult)
  {
    return info.GetReturnValue().Set(true);
  }
  OLETRACEOUT();
}

} // namespace node_win32ole
