#ifndef __CLIENT_H__
#define __CLIENT_H__

#include <node.h>
#include <nan.h>
#include "ole32core.h"

namespace node_win32ole {

class Client : public node::ObjectWrap {
public:
  static Nan::Persistent<v8::FunctionTemplate> clazz;
  static void Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target);
  static NAN_METHOD(New);
  static NAN_METHOD(Dispatch);
  static NAN_METHOD(Finalize);
public:
  Client() : finalized(false) {}
  ~Client() { if(!finalized) Finalize(); }
protected:
  void Finalize();
protected:
  bool finalized;
  ole32core::OLE32core oc;
};

} // namespace node_win32ole

#endif // __CLIENT_H__
