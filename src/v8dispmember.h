#ifndef __V8DISPMEMBER_H__
#define __V8DISPMEMBER_H__

#include <node.h>
#include <nan.h>
#include "node_win32ole.h"
#include "ole32core.h"

namespace node_win32ole {

class V8Dispatch;

// intended to be used as a fallback reference to an OLE property when we don't know if it's a property or a method
class V8DispMember : public node::ObjectWrap {
public:
  static Nan::Persistent<FunctionTemplate> clazz;
  static void Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target);
  static NAN_METHOD(OLEValue);
  static NAN_METHOD(OLEStringValue);
  static NAN_METHOD(OLELocaleStringValue);
  static NAN_METHOD(OLEPrimitiveValue);
  static Handle<Object> CreateNew(Handle<Object> dispatch, DISPID id); // *** private
  static NAN_METHOD(New);
  static NAN_METHOD(OLECall);
  static NAN_PROPERTY_GETTER(OLEGetAttr);
  static NAN_PROPERTY_SETTER(OLESetAttr);
  static NAN_INDEX_GETTER(OLEGetIdxAttr);
  static NAN_INDEX_SETTER(OLESetIdxAttr);
public:
  inline V8DispMember(DISPID id) : memberId(id) {}
protected:
  static Local<Value> resolveValue(Local<Object> thisObject);
  static Local<Value> resolveValueChain(Local<Object> thisObject, const char* prop);
  V8Dispatch* getDispatch(Local<Object> dispMemberObj);
  DISPID memberId;
};

} // namespace node_win32ole

#endif // __V8VARIANT_H__
