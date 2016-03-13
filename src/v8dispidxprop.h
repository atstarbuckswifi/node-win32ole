#ifndef __V8DISPIDXPROP_H__
#define __V8DISPIDXPROP_H__

#include <node.h>
#include <nan.h>
#include "node_win32ole.h"
#include "ole32core.h"

namespace node_win32ole {

class V8Dispatch;

// intended to be used as a fallback reference to an OLE property when we don't know if it's a property or a method
class V8DispIdxProperty : public node::ObjectWrap {
public:
  static Nan::Persistent<FunctionTemplate> clazz;
  static void Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target);
  static NAN_METHOD(OLEStringValue);
  static MaybeLocal<Object> CreateNew(Handle<Object> dispatch, Handle<String> property, DISPID id); // *** private
  static NAN_METHOD(New);
  static NAN_PROPERTY_GETTER(OLEGetAttr);
  static NAN_PROPERTY_SETTER(OLESetAttr);
  static NAN_INDEX_GETTER(OLEGetIdxAttr);
  static NAN_INDEX_SETTER(OLESetIdxAttr);
public:
  inline V8DispIdxProperty(const wchar_t* n, DISPID id) : memberId(id),name(n) {}
protected:
  V8Dispatch* getDispatch(Local<Object> dispMemberObj);
  DISPID memberId;
  std::wstring name;
};

} // namespace node_win32ole

#endif // __V8DISPIDXPROP_H__
