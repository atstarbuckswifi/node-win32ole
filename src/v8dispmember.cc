/*
  v8dispmember.cc
*/

#include "v8dispmember.h"
#include <node.h>
#include <nan.h>
#include "v8dispatch.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Nan::Persistent<FunctionTemplate> V8DispMember::clazz;

void V8DispMember::Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target)
{
  Nan::HandleScope scope;
  Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(1);
  t->SetClassName(Nan::New("V8DispMember").ToLocalChecked());
  Nan::SetPrototypeMethod(t, "valueOf", OLEPrimitiveValue);
  Nan::SetPrototypeMethod(t, "toString", OLEStringValue);
  Nan::SetPrototypeMethod(t, "toLocaleString", OLELocaleStringValue);
/*
 In ParseUnaryExpression() < v8/src/parser.cc >
 v8::Object::ToBoolean() is called directly for unary operator '!'
 instead of v8::Object::valueOf()
 so NamedPropertyHandler will not be called
 Local<Boolean> ToBoolean(); // How to fake ? override v8::Value::ToBoolean
*/
  Local<ObjectTemplate> instancetpl = t->InstanceTemplate();
  Nan::SetCallAsFunctionHandler(instancetpl, OLECall);
  Nan::SetNamedPropertyHandler(instancetpl, OLEGetAttr, OLESetAttr);
  Nan::SetIndexedPropertyHandler(instancetpl, OLEGetIdxAttr, OLESetIdxAttr);
  Nan::Set(target, Nan::New("V8DispMember").ToLocalChecked(), t->GetFunction());
  clazz.Reset(t);
}

Local<Value> V8DispMember::resolveValue(Local<Object> thisObject)
{
  V8DispMember *vThis = V8DispMember::Unwrap<V8DispMember>(thisObject);
  CHECK_V8_UNDEFINED(V8DispMember, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  if (!vDisp) return Nan::Undefined();
  return vDisp->OLEGet(vThis->memberId);
}

Local<Value> V8DispMember::resolveValueChain(Local<Object> thisObject, const char* prop)
{
  Local<Value> vResult = resolveValue(thisObject);
  if (vResult->IsUndefined()) return Nan::Undefined();
  vResult = INSTANCE_CALL(Nan::To<Object>(vResult).ToLocalChecked(), prop, 0, NULL);
  return vResult;
}

NAN_METHOD(V8DispMember::OLEValue)
{
  OLETRACEIN();
  Local<Value> vResult = resolveValue(info.This());
  if (!vResult->IsUndefined())
  {
    return info.GetReturnValue().Set(vResult);
  }
  OLETRACEOUT();
}

/**
 * Like OLEValue, but goes one step further and reduces IDispatch objects
 * to actual JS objects. This enables things like console.log().
 **/
NAN_METHOD(V8DispMember::OLEPrimitiveValue)
{
  OLETRACEIN();
  Local<Value> vResult = resolveValueChain(info.This(), "valueOf");
  if (vResult->IsUndefined()) return;
  OLETRACEOUT();
  return info.GetReturnValue().Set(vResult);
}

NAN_METHOD(V8DispMember::OLEStringValue)
{
  OLETRACEIN();
  Local<Value> vResult = resolveValueChain(info.This(), "toString");
  if (vResult->IsUndefined()) return;
  OLETRACEOUT();
  return info.GetReturnValue().Set(vResult);
}

NAN_METHOD(V8DispMember::OLELocaleStringValue)
{
  OLETRACEIN();
  Local<Value> vResult = resolveValueChain(info.This(), "toLocaleString");
  if (vResult->IsUndefined()) return;
  OLETRACEOUT();
  return info.GetReturnValue().Set(vResult);
}

MaybeLocal<Object> V8DispMember::CreateNew(Handle<Object> dispatch, DISPID id)
{
  DISPFUNCIN();
  Local<FunctionTemplate> localClazz = Nan::New(clazz);
  Local<v8::Value> args[] = { dispatch, Nan::New<Int32>(id) };
  int argc = sizeof(args) / sizeof(args[0]); // == 2
  return Nan::NewInstance(Nan::GetFunction(localClazz).ToLocalChecked(), argc, args);
  DISPFUNCOUT();
}

NAN_METHOD(V8DispMember::New)
{
  DISPFUNCIN();
  if(!info.IsConstructCall())
    return Nan::ThrowTypeError("Use the new operator to create new V8DispMember objects");
  if (info.Length() != 2)
    return Nan::ThrowTypeError("Must be constructed with (V8Dispatch, DISPID)");
  Local<FunctionTemplate> v8DispatchClazz = Nan::New(V8Dispatch::clazz);
  if (!info[0]->IsObject() || !v8DispatchClazz->HasInstance(info[0]))
    return Nan::ThrowTypeError("First argument must be a V8Dispatch object");
  if(!info[1]->IsInt32())
    return Nan::ThrowTypeError("Second argument must be a DISPID");
  Nan::MaybeLocal<Int32> maybeDispId = Nan::To<Int32>(info[1]);
  if(maybeDispId.IsEmpty())
    return Nan::ThrowTypeError("Second argument must be a DISPID");
  DISPID id = maybeDispId.ToLocalChecked()->Int32Value();
  V8DispMember *v = new V8DispMember(id); // must catch exception
  CHECK_V8(V8DispMember, v);
  Local<Object> thisObject = info.This();
  v->Wrap(thisObject); // InternalField[0]
  Nan::ForceSet(thisObject, Nan::New<String>("_dispatch").ToLocalChecked(), info[0]);
  DISPFUNCOUT();
  return info.GetReturnValue().Set(thisObject);
}

V8Dispatch* V8DispMember::getDispatch(Local<Object> dispMemberObj)
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

NAN_METHOD(V8DispMember::OLECall)
{
  OLETRACEIN();
  OLETRACEARGS();
  Local<Object> thisObject = info.This();
  V8DispMember *vThis = V8DispMember::Unwrap<V8DispMember>(thisObject);
  CHECK_V8(V8DispMember, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);
  Local<Value> vResult = vDisp->OLECall(vThis->memberId, info);
  if (!vResult->IsUndefined())
  {
    return info.GetReturnValue().Set(vResult);
  }
  OLETRACEOUT();
}

NAN_PROPERTY_GETTER(V8DispMember::OLEGetAttr)
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

  String::Value vProperty(property);
  if (wcscmp((const wchar_t*)*vProperty, L"_") == 0)
  {
    Local<Value> vResult = resolveValue(thisObject);
	if (vResult->IsUndefined()) return; // exception?
    return info.GetReturnValue().Set(vResult);
  }

  // Resolve ourselves as a property reference
  Local<Value> vResult = resolveValue(thisObject);
  if (vResult->IsUndefined()) return;

  // try to pass the request to the resulting value
  if(!vResult->IsObject())
  {
    return Nan::ThrowTypeError("Value of this property is not an object");
  }
  Local<Object> hResult = Local<Object>::Cast(vResult);
  Nan::MaybeLocal<Value> mResult = Nan::Get(hResult, property);
  if (mResult.IsEmpty()) return;
  vResult = mResult.ToLocalChecked();
  if (vResult->IsUndefined()) return;
  OLETRACEOUT();
  return info.GetReturnValue().Set(vResult);
}

NAN_PROPERTY_SETTER(V8DispMember::OLESetAttr)
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
  V8DispMember *vThis = V8DispMember::Unwrap<V8DispMember>(thisObject);
  CHECK_V8(V8DispMember, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);
  Local<Value> vResult = vDisp->OLEGet(vThis->memberId);
  if (vResult->IsUndefined()) return;

  // try to pass the request to the resulting value
  if (!vResult->IsObject())
  {
    return Nan::ThrowTypeError("Value of this property is not an object");
  }
  Local<Object> hResult = Local<Object>::Cast(vResult);
  Nan::Maybe<bool> mbResult = Nan::Set(hResult, property, value);
  if (mbResult.IsNothing()) return;
  OLETRACEOUT();
  return info.GetReturnValue().Set(mbResult.IsJust());
}

NAN_INDEX_GETTER(V8DispMember::OLEGetIdxAttr)
{
  OLETRACEIN();
  {
    OLETRACEPREARGV(Nan::New(index));
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();

  // Resolve ourselves as an indexed property reference
  V8DispMember *vThis = V8DispMember::Unwrap<V8DispMember>(thisObject);
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

NAN_INDEX_SETTER(V8DispMember::OLESetIdxAttr)
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
  V8DispMember *vThis = V8DispMember::Unwrap<V8DispMember>(thisObject);
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
