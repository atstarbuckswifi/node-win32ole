#ifndef __CLIENT_H__
#define __CLIENT_H__

#include <node.h>
#include <nan.h>
#include "node_win32ole.h"

using namespace v8;
namespace ole32core {
  class OLE32core;
}

namespace node_win32ole {

class Client : public node::ObjectWrap {
public:
  static Nan::Persistent<FunctionTemplate> clazz;
  static void Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target);
  static NAN_METHOD(New);
  static NAN_METHOD(Dispatch);
  static NAN_METHOD(Finalize);
public:
  Client() : node::ObjectWrap(), finalized(false) {}
  ~Client() { if(!finalized) Finalize(); }
protected:
  static void Dispose(const Nan::WeakCallbackInfo<ole32core::OLE32core> &data);
  void Finalize();
protected:
  bool finalized;
};

} // namespace node_win32ole

#endif // __CLIENT_H__
