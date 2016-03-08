#ifndef __V8VARIANT_H__
#define __V8VARIANT_H__

#include <node.h>
#include <nan.h>
#include "node_win32ole.h"
#include "ole32core.h"

namespace node_win32ole {

extern Handle<Value> NewOleException(HRESULT hr);
extern Handle<Value> NewOleException(HRESULT hr, const ole32core::ErrorInfo& info);

class V8Variant : public node::ObjectWrap {
public:
  static Nan::Persistent<FunctionTemplate> clazz;
  static void Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target);
  static NAN_METHOD(OLEIsA);
  static NAN_METHOD(OLEVTName);
  static NAN_METHOD(OLEBoolean); // *** p.
  static NAN_METHOD(OLEInt32); // *** p.
  static NAN_METHOD(OLENumber); // *** p.
  static NAN_METHOD(OLEDate); // *** p.
  static NAN_METHOD(OLEUtf8); // *** p.
  static NAN_METHOD(OLEValue);
  static NAN_METHOD(OLEStringValue);
  static NAN_METHOD(OLELocaleStringValue);
  static Handle<Object> CreateUndefined(void); // *** private
  static NAN_METHOD(New);
  static NAN_METHOD(Finalize);
  static Local<Value> VariantToValue(const VARIANT& ocv);
  static ole32core::OCVariant *ValueToVariant(Handle<Value> v); // *** private
public:
  V8Variant() : finalized(false) {}
  ~V8Variant() { if(!finalized) Finalize(); }
  ole32core::OCVariant ocv;
protected:
  void Finalize();
  static Local<Value> resolveValueChain(Local<Object> thisObject, const char* prop);
  static Local<Date> OLEDateToObject(const DATE& dt);
  static Local<Value> ArrayToValue(const SAFEARRAY& a);
  static Local<Value> ArrayToValueSlow(const SAFEARRAY& a, VARTYPE vt, LONG* idices, unsigned numIdx);
  static Local<Value> ArrayPrimitiveToValue(void* loc, VARTYPE vt, unsigned cbElements, unsigned idx);
protected:
  bool finalized;
};

} // namespace node_win32ole

#endif // __V8VARIANT_H__
