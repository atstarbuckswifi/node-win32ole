#ifndef __V8VARIANT_H__
#define __V8VARIANT_H__

#include <node.h>
#include <nan.h>
#include "node_win32ole.h"
#include "ole32core.h"

namespace node_win32ole {

typedef struct _fundamental_attr {
  bool obsoleted;
  const char *name;
  NAN_METHOD((*func));
} fundamental_attr;

extern Handle<Value> NewOleException(HRESULT hr, const std::wstring& msg = L"");

class V8Variant : public node::ObjectWrap {
public:
  static Nan::Persistent<FunctionTemplate> clazz;
  static void Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target);
  static ole32core::OCVariant *CreateOCVariant(Handle<Value> v); // *** private
  static NAN_METHOD(OLEIsA);
  static NAN_METHOD(OLEVTName);
  static NAN_METHOD(OLEBoolean); // *** p.
  static NAN_METHOD(OLEInt32); // *** p.
  static NAN_METHOD(OLENumber); // *** p.
  static NAN_METHOD(OLEDate); // *** p.
  static NAN_METHOD(OLEUtf8); // *** p.
  static NAN_METHOD(OLEValue);
  static NAN_METHOD(OLEPrimitiveValue);
  static Handle<Object> CreateUndefined(void); // *** private
  static NAN_METHOD(New);
  static Handle<Value> OLEFlushCarryOver(Handle<Value> v); // *** p.
  static NAN_METHOD(OLECall);
  static NAN_METHOD(OLEGet);
  static NAN_METHOD(OLESet);
  static NAN_METHOD(OLECallComplete); // *** p.
  static NAN_PROPERTY_GETTER(OLEGetAttr);
  static NAN_PROPERTY_SETTER(OLESetAttr);
  static NAN_METHOD(Finalize);
public:
  V8Variant() : finalized(false), property_carryover() {}
  ~V8Variant() { if(!finalized) Finalize(); }
  ole32core::OCVariant ocv;
protected:
  static void Dispose(const Nan::WeakCallbackInfo<ole32core::OCVariant> &data);
  void Finalize();
  static Local<Date> OLEDateToObject(const DATE& dt);
  static Local<Value> VariantToValue(Handle<Object> thisObject, const VARIANT& ocv);
  static Local<Value> ArrayToValue(Handle<Object> thisObject, const SAFEARRAY& a);
  static Local<Value> ArrayToValueSlow(Handle<Object> thisObject, const SAFEARRAY& a, VARTYPE vt, LONG* idices, unsigned numIdx);
  static Local<Value> ArrayPrimitiveToValue(Handle<Object> thisObject, void* loc, VARTYPE vt, unsigned cbElements, unsigned idx);
protected:
  bool finalized;
  std::string property_carryover;
};

} // namespace node_win32ole

#endif // __V8VARIANT_H__
