#ifndef __V8DISPATCH_H__
#define __V8DISPATCH_H__

#include <functional>
#include <map>
#include <nan.h>
#include <node.h>
#include "node_win32ole.h"
#include "ole32core.h"

namespace node_win32ole {

class V8Dispatch : public node::ObjectWrap {
public:
  static Nan::Persistent<FunctionTemplate> clazz;
  static void Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target);
  static NAN_METHOD(OLEValue);
  static NAN_METHOD(OLEPrimitiveValue);
  static NAN_METHOD(OLEStringValue);
  static NAN_METHOD(OLELocaleStringValue);
  static MaybeLocal<Object> CreateNew(IDispatch* disp); // *** private
  static NAN_METHOD(New);
  static NAN_PROPERTY_GETTER(OLEGetAttr);
  static NAN_PROPERTY_SETTER(OLESetAttr);
  static NAN_PROPERTY_ENUMERATOR(OLEEnumAttr);
  static NAN_PROPERTY_QUERY(OLEQueryAttr);
  static NAN_INDEX_GETTER(OLEGetIdxAttr);
  static NAN_INDEX_SETTER(OLESetIdxAttr);
  static NAN_METHOD(Finalize);
public:
  V8Dispatch() : finalized(false) {}
  ~V8Dispatch() { if(!finalized) Finalize(); }
  ole32core::OCDispatch ocd;

  Local<Value> OLECall(DISPID propID, Nan::NAN_METHOD_ARGS_TYPE info, WORD targetType = DISPATCH_METHOD | DISPATCH_PROPERTYGET);
  Local<Value> OLECall(DISPID propID, int argc = 0, Local<Value> argv[] = NULL, WORD targetType = DISPATCH_METHOD | DISPATCH_PROPERTYGET);
  Local<Value> OLEGet(DISPID propID, int argc = 0, Local<Value> argv[] = NULL);
  bool OLESet(DISPID propID, int argc = 0, Local<Value> argv[] = NULL);

public:
  std::wstring m_typeName;

protected:
  static Local<Value> resolveValueChain(Local<Object> thisObject, const char* prop);
  HRESULT interrogateType();
  void Finalize();
  bool finalized;

  enum EMemberAttr
  {
    ma_IsProperty = 1,
    ma_IsIndexedProperty = 2,
    ma_IsReadOnly = 4,
    ma_IsHidden = 8
  };

  struct MemberInfo
  {
    DISPID memberID;
    int attrs;
  };

  struct CaseInsensitive : public std::binary_function<std::wstring, std::wstring, bool>
  {
    bool operator()(const std::wstring& left, const std::wstring& right) const
    {
      return wcsicmp(left.c_str(), right.c_str()) < 0;
    }
  };

  typedef std::map<std::wstring, MemberInfo, CaseInsensitive> TMemberMap;
  bool m_bHasDefaultProp;
  TMemberMap m_members;
};

} // namespace node_win32ole

#endif // __V8DISPATCH_H__
