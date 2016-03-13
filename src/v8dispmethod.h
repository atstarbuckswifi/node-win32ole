#ifndef __V8DISPMETHOD_H__
#define __V8DISPMETHOD_H__

#include <node.h>
#include <nan.h>
#include "node_win32ole.h"
#include "ole32core.h"

namespace node_win32ole {

class V8Dispatch;

// Represents an OLE method to be called
class V8DispMethod : public node::ObjectWrap {
public:
  static Nan::Persistent<FunctionTemplate> clazz;
  static void Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target);
  static NAN_METHOD(OLEStringValue);
  static MaybeLocal<Object> CreateNew(Handle<Object> dispatch, WORD targetType, Handle<String> property, DISPID id); // *** private
  static NAN_METHOD(New);
  static NAN_METHOD(OLECall);
public:
  inline V8DispMethod(WORD type, const wchar_t* n, DISPID id) :targetType(type), name(n), memberId(id) {}
protected:
  V8Dispatch* getDispatch(Local<Object> dispMethodObj);
  WORD targetType;
  std::wstring name;
  DISPID memberId;
};

} // namespace node_win32ole

#endif // __V8DISPMETHOD_H__
