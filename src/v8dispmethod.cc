/*
  v8dispmethod.cc
*/

#include "v8dispmethod.h"
#include <node.h>
#include <nan.h>
#include "v8dispatch.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Nan::Persistent<FunctionTemplate> V8DispMethod::clazz;

void V8DispMethod::Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target)
{
  Nan::HandleScope scope;
  Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(1);
  t->SetClassName(Nan::New("V8DispMethod").ToLocalChecked());
  Nan::SetPrototypeMethod(t, "valueOf", OLEStringValue);
  Nan::SetPrototypeMethod(t, "toString", OLEStringValue);
  Nan::SetPrototypeMethod(t, "toLocaleString", OLEStringValue);

  Local<ObjectTemplate> instancetpl = t->InstanceTemplate();
  Nan::SetCallAsFunctionHandler(instancetpl, OLECall);
  Nan::Set(target, Nan::New("V8DispMethod").ToLocalChecked(), t->GetFunction());
  clazz.Reset(t);
}

NAN_METHOD(V8DispMethod::OLEStringValue)
{
  OLETRACEIN();
  Local<Object> thisObject = info.This();
  V8DispMethod *vThis = V8DispMethod::Unwrap<V8DispMethod>(thisObject);
  CHECK_V8(V8DispMethod, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);
  std::wstring fullName = (vDisp->m_typeName.empty() ? L"Object" : vDisp->m_typeName) + L"." + vThis->name;
  OLETRACEOUT();
  return info.GetReturnValue().Set(Nan::New((const uint16_t*)fullName.c_str()).ToLocalChecked());
}

MaybeLocal<Object> V8DispMethod::CreateNew(Handle<Object> dispatch, WORD targetType, Handle<String> property, DISPID id) // *** private
{
  DISPFUNCIN();
  Local<FunctionTemplate> localClazz = Nan::New(clazz);
  Local<v8::Value> args[] = { dispatch, Nan::New<Int32>(targetType), property, Nan::New<Int32>(id) };
  int argc = sizeof(args) / sizeof(args[0]); // == 4
  return Nan::NewInstance(Nan::GetFunction(localClazz).ToLocalChecked(), argc, args);
  DISPFUNCOUT();
}

NAN_METHOD(V8DispMethod::New)
{
  DISPFUNCIN();
  if(!info.IsConstructCall())
    return Nan::ThrowTypeError("Use the new operator to create new V8DispMethod objects");
  if (info.Length() != 4)
    return Nan::ThrowTypeError("Must be constructed with (V8Dispatch, propType, propName, DISPID)");
  Local<FunctionTemplate> v8DispatchClazz = Nan::New(V8Dispatch::clazz);
  if (!info[0]->IsObject() || !v8DispatchClazz->HasInstance(info[0]))
    return Nan::ThrowTypeError("First argument must be a V8Dispatch object");
  if (!info[1]->IsInt32())
    return Nan::ThrowTypeError("Second argument must be a propType");
  Nan::MaybeLocal<Int32> maybeType = Nan::To<Int32>(info[1]);
  if(!info[2]->IsString() && !info[2]->IsStringObject())
    return Nan::ThrowTypeError("Third argument must be a String");
  Handle<String> hName = Nan::To<String>(info[2]).ToLocalChecked();
  if (!info[3]->IsInt32())
    return Nan::ThrowTypeError("Fourth argument must be a DISPID");
  Nan::MaybeLocal<Int32> maybeDispId = Nan::To<Int32>(info[3]);

  WORD targetType = maybeType.ToLocalChecked()->Int32Value();
  String::Value vName(hName);
  DISPID id = maybeDispId.ToLocalChecked()->Int32Value();

  V8DispMethod *v = new V8DispMethod(targetType, (const wchar_t*)*vName, id); // must catch exception
  CHECK_V8(V8DispMember, v);
  Local<Object> thisObject = info.This();
  v->Wrap(thisObject); // InternalField[0]
  Nan::ForceSet(thisObject, Nan::New<String>("_dispatch").ToLocalChecked(), info[0]);
  DISPFUNCOUT();
  return info.GetReturnValue().Set(thisObject);
}

V8Dispatch* V8DispMethod::getDispatch(Local<Object> dispMemberObj)
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

NAN_METHOD(V8DispMethod::OLECall)
{
  OLETRACEIN();
  OLETRACEARGS();
  Local<Object> thisObject = info.This();
  V8DispMethod *vThis = V8DispMethod::Unwrap<V8DispMethod>(thisObject);
  CHECK_V8(V8DispMember, vThis);
  V8Dispatch* vDisp = vThis->getDispatch(thisObject);
  CHECK_V8(V8Dispatch, vDisp);
  Local<Value> vResult = vDisp->OLECall(vThis->memberId, info, vThis->targetType);
  if (!vResult->IsUndefined())
  {
    return info.GetReturnValue().Set(vResult);
  }
  OLETRACEOUT();
}

} // namespace node_win32ole
