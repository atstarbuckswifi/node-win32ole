/*
  v8dispatch.cc
*/

#include "v8dispatch.h"
#include <node.h>
#include <nan.h>
#include <set>
#include "v8dispidxprop.h"
#include "v8dispmember.h"
#include "v8dispmethod.h"
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
  Nan::SetNamedPropertyHandler(instancetpl, OLEGetAttr, OLESetAttr, OLEQueryAttr, NULL, OLEEnumAttr);
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

  HRESULT hr = vThis->interrogateType();
  if (FAILED(hr)) return Nan::ThrowError(NewOleException(hr));

  if (vThis->m_members.empty() || vThis->m_bHasDefaultProp)
  {
    // we either have no type information or we know there is a default property here, fetch it
    Local<Value> vResult = vThis->OLEGet(DISPID_VALUE);
    if (vResult->IsUndefined()) info.GetReturnValue().Set(vResult);
    return;
  }

  // we have no special action, return our object type name
  OLETRACEOUT();
  return info.GetReturnValue().Set(Nan::New((const uint16_t*)vThis->m_typeName.c_str()).ToLocalChecked());
}

Local<Value> V8Dispatch::resolveValueChain(Local<Object> thisObject, const char* prop)
{
  OLETRACEIN();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(thisObject);
  CHECK_V8_UNDEFINED(V8Dispatch, vThis);

  HRESULT hr = vThis->interrogateType();
  if (FAILED(hr))
  {
    Nan::ThrowError(NewOleException(hr));
    return Nan::Undefined();
  }

  if (vThis->m_members.empty() || vThis->m_bHasDefaultProp)
  {
    // we either have no type information or we know there is a default property here, fetch it
    Local<Value> vResult = vThis->OLEGet(DISPID_VALUE);
    if (vResult->IsUndefined()) return Nan::Undefined();
    vResult = INSTANCE_CALL(Nan::To<Object>(vResult).ToLocalChecked(), prop, 0, NULL);
    return vResult;
  }

  // we have no special action, return our object type name
  OLETRACEOUT();
  return Nan::New((const uint16_t*)vThis->m_typeName.c_str()).ToLocalChecked();
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

MaybeLocal<Object> V8Dispatch::CreateNew(IDispatch* disp)
{
  DISPFUNCIN();
  Local<FunctionTemplate> localClazz = Nan::New(clazz);
  MaybeLocal<Object> mInstance = Nan::NewInstance(Nan::GetFunction(localClazz).ToLocalChecked(), 0, NULL);
  if (mInstance.IsEmpty()) return mInstance;
  Local<Object> instance = mInstance.ToLocalChecked();
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
  return info.GetReturnValue().Set(thisObject);
}

Local<Value> V8Dispatch::OLECall(DISPID propID, Nan::NAN_METHOD_ARGS_TYPE info, WORD targetType /* = DISPATCH_METHOD | DISPATCH_PROPERTYGET */)
{
  DISPFUNCIN();
  int argc = info.Length();
  Local<Value>* argv = (Local<Value>*)alloca(sizeof(Local<Value>) * argc);
  for (int idx = 0; idx < argc; ++idx)
  {
    *(new(argv + idx) Local<Value>) = info[idx];
  }
  Local<Value> hResult = OLECall(propID, argc, argv, targetType);
  for (int idx = 0; idx < argc; ++idx)
  {
    (argv + idx)->~Local<Value>();
  }
  DISPFUNCOUT();
  return hResult;
}

Local<Value> V8Dispatch::OLECall(DISPID propID, int argc, Local<Value> argv[], WORD targetType /* = DISPATCH_METHOD | DISPATCH_PROPERTYGET */ )
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
  HRESULT hr = ocd.invoke(targetType, propID, &rv.v, errInfo, argc, argchain); // argchain will be deleted automatically
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
  HRESULT hr = ocd.invoke(DISPATCH_PROPERTYGET, propID, &rv.v, errInfo, argc, argchain); // argchain will be deleted automatically
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
  HRESULT hr = ocd.invoke(DISPATCH_PROPERTYPUT, propID, NULL, errInfo, argc, argchain); // argchain will be deleted automatically
  if (FAILED(hr))
  {
    Nan::ThrowError(NewOleException(hr, errInfo));
    return false;
  }
  OLETRACEOUT();
  return true;
}

HRESULT V8Dispatch::interrogateType()
{
  if (!m_members.empty() || !ocd.disp) return S_FALSE;
  ITypeInfo* tinfo = ocd.getTypeInfo();

  // pull this object type
  BSTR bTypeName;
  HRESULT hr = tinfo->GetDocumentation(MEMBERID_NIL, &bTypeName, NULL, NULL, NULL);
  if (SUCCEEDED(hr))
  {
    m_typeName = std::wstring(bTypeName, SysStringLen(bTypeName));
    SysFreeString(bTypeName);
  }

  TYPEATTR* tattr;
  hr = tinfo->GetTypeAttr(&tattr);
  if (FAILED(hr)) return hr; // can't really recover from this one

  m_bHasDefaultProp = false;
  m_members.clear();
  
  BSTR bMemName;
  UINT numNames;

  // pull functions and properties
  for (int i = 0; i < tattr->cFuncs; ++i) {
    FUNCDESC *funcdesc;
    hr = tinfo->GetFuncDesc(i, &funcdesc);
    if(SUCCEEDED(hr))
    {
      if (funcdesc->memid == DISPID_VALUE) m_bHasDefaultProp = true;
      if(!(funcdesc->wFuncFlags & FUNCFLAG_FRESTRICTED))
      {
        hr = tinfo->GetNames(funcdesc->memid, &bMemName, 1, &numNames);
        if (SUCCEEDED(hr))
        {
          std::wstring memberName(bMemName, SysStringLen(bMemName));
          SysFreeString(bMemName);
          TMemberMap::iterator lookup = m_members.find(memberName);
          if (lookup == m_members.end())
          {
            lookup = m_members.insert(TMemberMap::value_type(memberName, MemberInfo())).first;
            MemberInfo& info = lookup->second;
            info.memberID = funcdesc->memid;
            info.attrs = ma_IsReadOnly;
            if (funcdesc->invkind != INVOKE_FUNC) info.attrs |= ma_IsProperty;
            if (funcdesc->wFuncFlags & (FUNCFLAG_FHIDDEN | FUNCFLAG_FNONBROWSABLE)) info.attrs |= ma_IsHidden;
          }
          MemberInfo& info = lookup->second;
          if (info.memberID == funcdesc->memid)
          {
            if (funcdesc->invkind == INVOKE_PROPERTYPUT || funcdesc->invkind == INVOKE_PROPERTYPUT) info.attrs &= ma_IsReadOnly;
            if (funcdesc->invkind == INVOKE_PROPERTYGET && funcdesc->cParams) info.attrs |= ma_IsIndexedProperty;
          }
        }
      }
      tinfo->ReleaseFuncDesc(funcdesc);
    }
  }

  for (int i = 0; i < tattr->cVars; ++i) {
    VARDESC *vardesc;
    hr = tinfo->GetVarDesc(i, &vardesc);
    if (SUCCEEDED(hr))
    {
      if (vardesc->memid == DISPID_VALUE) m_bHasDefaultProp = true;
      if(!(vardesc->wVarFlags & VARFLAG_FRESTRICTED))
      {
        hr = tinfo->GetNames(vardesc->memid, &bMemName, 1, &numNames);
        if (SUCCEEDED(hr))
        {
          std::wstring memberName(bMemName, SysStringLen(bMemName));
          SysFreeString(bMemName);
          if(m_members.find(memberName) == m_members.end())
          {
            MemberInfo& info = m_members.insert(TMemberMap::value_type(memberName, MemberInfo())).first->second;
            info.memberID = vardesc->memid;
            info.attrs = ma_IsProperty;
            if (vardesc->wVarFlags & VARFLAG_FREADONLY) info.attrs |= ma_IsReadOnly;
            if (vardesc->wVarFlags & (VARFLAG_FHIDDEN | VARFLAG_FNONBROWSABLE)) info.attrs |= ma_IsHidden;
          }
        }
      }
      tinfo->ReleaseVarDesc(vardesc);
    }
  }

  tinfo->ReleaseTypeAttr(tattr);
  return S_OK;
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
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(thisObject);
  CHECK_V8(V8Dispatch, vThis);

  HRESULT hr = vThis->interrogateType();
  if (FAILED(hr)) return Nan::ThrowError(NewOleException(hr));

  String::Value vProperty(property);
  const wchar_t* szProperty = (const wchar_t*)*vProperty;
  if (wcscmp(szProperty, L"_") == 0 && (vThis->m_members.empty() || vThis->m_bHasDefaultProp))
  {
    Local<Value> vResult = vThis->OLEGet(DISPID_VALUE);
    if (vResult->IsUndefined()) return; // exception?
    return info.GetReturnValue().Set(vResult);
  }

  // try to resolve this as a member of our object
  if (vThis->m_members.empty())
  { // no type information, we're running blind
    BSTR bName = ::SysAllocString(szProperty);
    if (bName)
    {
      DISPID dispID;
      HRESULT hr = vThis->ocd.disp->GetIDsOfNames(IID_NULL, &bName, 1, LOCALE_USER_DEFAULT, &dispID);
      ::SysFreeString(bName);
      if (SUCCEEDED(hr))
      {
        MaybeLocal<Object> vDispMember = V8DispMember::CreateNew(thisObject, dispID);
        if(!vDispMember.IsEmpty()) info.GetReturnValue().Set(vDispMember.ToLocalChecked());
        return;
      }
    }
  }
  else
  {
    TMemberMap::const_iterator lookup = vThis->m_members.find(szProperty);
    if (lookup != vThis->m_members.end())
    {
      const std::wstring& memberName = lookup->first;
      const MemberInfo& minfo = lookup->second;
      if (minfo.attrs & ma_IsProperty)
      {
        if (minfo.attrs & ma_IsIndexedProperty)
        {
          // create indexed property object here
          MaybeLocal<Object> vDispIdxProp = V8DispIdxProperty::CreateNew(thisObject, Nan::New((const uint16_t*)memberName.c_str()).ToLocalChecked(), minfo.memberID);
          if(!vDispIdxProp.IsEmpty()) info.GetReturnValue().Set(vDispIdxProp.ToLocalChecked());
          return;
        } else {
          // fetch property value now
          Local<Value> vResult = vThis->OLEGet(minfo.memberID);
          if (vResult->IsUndefined()) return; // exception?
          return info.GetReturnValue().Set(vResult);
        }
      } else {
        // create method property object here
        MaybeLocal<Object> vDispMethod = V8DispMethod::CreateNew(thisObject, DISPATCH_METHOD, Nan::New((const uint16_t*)memberName.c_str()).ToLocalChecked(), minfo.memberID);
        if(!vDispMethod.IsEmpty()) info.GetReturnValue().Set(vDispMethod.ToLocalChecked());
        return;
      }
    }
    if(wcsnicmp(szProperty, L"get_", 4) == 0)
    {
      TMemberMap::const_iterator lookup = vThis->m_members.find(szProperty + 4);
      if (lookup != vThis->m_members.end())
      {
        const std::wstring& memberName = lookup->first;
        const MemberInfo& minfo = lookup->second;
        if (minfo.attrs & ma_IsProperty)
        {
          // create method property object here
          MaybeLocal<Object> vDispMethod = V8DispMethod::CreateNew(thisObject, DISPATCH_PROPERTYGET, Nan::New((const uint16_t*)memberName.c_str()).ToLocalChecked(), minfo.memberID);
          if(!vDispMethod.IsEmpty()) info.GetReturnValue().Set(vDispMethod.ToLocalChecked());
          return;
        }
      }
    }
    if (wcsnicmp(szProperty, L"put_", 4) == 0)
    {
      TMemberMap::const_iterator lookup = vThis->m_members.find(szProperty + 4);
      if (lookup != vThis->m_members.end())
      {
        const std::wstring& memberName = lookup->first;
        const MemberInfo& minfo = lookup->second;
        if (minfo.attrs & ma_IsProperty)
        {
          // create method property object here
          MaybeLocal<Object> vDispMethod = V8DispMethod::CreateNew(thisObject, DISPATCH_PROPERTYPUT, Nan::New((const uint16_t*)memberName.c_str()).ToLocalChecked(), minfo.memberID);
          if(!vDispMethod.IsEmpty()) info.GetReturnValue().Set(vDispMethod.ToLocalChecked());
          return;
        }
      }
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
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(thisObject);
  CHECK_V8(V8Dispatch, vThis);

  String::Value vProperty(property);
  if (wcscmp((const wchar_t*)*vProperty, L"_") == 0 && (vThis->m_members.empty() || vThis->m_bHasDefaultProp))
  {
    bool bResult = vThis->OLESet(DISPID_VALUE, 1, &value);
    return info.GetReturnValue().Set(bResult);
  }

  // try to resolve this as a member of our object
  if (vThis->m_members.empty())
  { // no type information, we're running blind
    BSTR bName = ::SysAllocString((const wchar_t*)*vProperty);
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
  }
  else
  {
    TMemberMap::const_iterator lookup = vThis->m_members.find((const wchar_t*)*vProperty);
    if (lookup != vThis->m_members.end())
    {
      const MemberInfo& minfo = lookup->second;
      if (minfo.attrs & ma_IsProperty)
      {
        if (minfo.attrs & ma_IsIndexedProperty)
        {
          return Nan::ThrowError("Cannot set this property without an index");
        } else {
          // set property value now
          bool bResult = vThis->OLESet(minfo.memberID, 1, &value);
          return info.GetReturnValue().Set(bResult);
        }
      } else {
        return Nan::ThrowError("Cannot assign a value to a method");
      }
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

NAN_PROPERTY_ENUMERATOR(V8Dispatch::OLEEnumAttr)
{
  OLETRACEIN();
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(thisObject);
  CHECK_V8(V8Dispatch, vThis);

  HRESULT hr = vThis->interrogateType();
  if (FAILED(hr)) return Nan::ThrowError(NewOleException(hr));

  typedef std::set<std::wstring> TStringSet;
  TStringSet props;

  if (vThis->m_members.empty() || vThis->m_bHasDefaultProp)
  {
    props.insert(L"_");
  }

  for (TMemberMap::const_iterator trans = vThis->m_members.begin(); trans != vThis->m_members.end(); trans++)
  {
    props.insert(trans->first);
  }

  if (!props.empty())
  {
    Local<Array> hReturn = Nan::New<Array>((int)props.size());
    int idx = 0;
    TStringSet::const_iterator trans = props.begin();
    for (; trans != props.end(); idx++, trans++)
    {
      Nan::Set(hReturn, idx, Nan::New<String>((const uint16_t*)trans->c_str()).ToLocalChecked());
    }
    info.GetReturnValue().Set(hReturn);
  }

  OLETRACEOUT();
}

NAN_PROPERTY_QUERY(V8Dispatch::OLEQueryAttr)
{
  OLETRACEIN();
  {
    OLETRACEPREARGV(property);
    OLETRACEARGV();
  }
  OLETRACEFLUSH();
  Local<Object> thisObject = info.This();
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(thisObject);
  CHECK_V8(V8Dispatch, vThis);

  HRESULT hr = vThis->interrogateType();
  if (FAILED(hr)) return Nan::ThrowError(NewOleException(hr));

  typedef std::set<std::wstring> TStringSet;
  TStringSet props;

  String::Value vProperty(property);
  if (wcscmp((const wchar_t*)*vProperty, L"_") == 0 && (vThis->m_members.empty() || vThis->m_bHasDefaultProp))
  { // TODO: is the default property r/o ?
    return info.GetReturnValue().Set(PropertyAttribute::DontDelete | PropertyAttribute::DontEnum);
  }

  // try to resolve this as a member of our object
  TMemberMap::const_iterator lookup = vThis->m_members.find((const wchar_t*)*vProperty);
  if (lookup != vThis->m_members.end())
  {
    const MemberInfo& minfo = lookup->second;
    return info.GetReturnValue().Set(
      PropertyAttribute::DontDelete |
      (minfo.attrs & ma_IsReadOnly ? PropertyAttribute::ReadOnly : 0) |
      (minfo.attrs & ma_IsHidden ? PropertyAttribute::DontEnum : 0));
  }

  OLETRACEOUT();
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
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(thisObject);
  CHECK_V8(V8Dispatch, vThis);

  HRESULT hr = vThis->interrogateType();
  if (FAILED(hr)) return Nan::ThrowError(NewOleException(hr));

  if (vThis->m_members.empty() || vThis->m_bHasDefaultProp)
  {
    // we either have no type information or we know there is a default property here, let's assume it's indexed
    Handle<Value> argv[] = { Nan::New(index) };
    int argc = sizeof(argv) / sizeof(argv[0]); // == 1
    Local<Value> vResult = vThis->OLEGet(DISPID_VALUE, argc, argv);
    if (!vResult->IsUndefined())
    {
      return info.GetReturnValue().Set(vResult);
    }
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
  V8Dispatch *vThis = V8Dispatch::Unwrap<V8Dispatch>(thisObject);
  CHECK_V8(V8Dispatch, vThis);

  HRESULT hr = vThis->interrogateType();
  if (FAILED(hr)) return Nan::ThrowError(NewOleException(hr));

  if (vThis->m_members.empty() || vThis->m_bHasDefaultProp)
  {
    // we either have no type information or we know there is a default property here, let's assume it's indexed
    Handle<Value> argv[] = { Nan::New(index), value };
    int argc = sizeof(argv) / sizeof(argv[0]); // == 1
    bool bResult = vThis->OLESet(DISPID_VALUE, argc, argv);
    if (bResult)
    {
      return info.GetReturnValue().Set(true);
    }
  }

  // we don't want the user creating arbitrary expandos
  OLETRACEOUT();
  return Nan::ThrowError("Cannot assign new properties to this object");
}

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
