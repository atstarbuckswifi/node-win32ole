/*
  node_win32ole.cc
  $ cp ~/.node-gyp/0.8.18/ia32/node.lib to ~/.node-gyp/0.8.18/node.lib
  $ node-gyp configure
  $ node-gyp build
  $ node test/init_win32ole.test.js
*/

#include "node_win32ole.h"
#include "client.h"
#include "v8variant.h"

using namespace v8;

namespace node_win32ole {

Nan::Persistent<Object> module_target;

NAN_METHOD(Method_version)
{
//  Nan::HandleScope scope; -- should be implicit in method calls
  Handle<Object> local = Nan::New<Object>(module_target);
  auto ver = Nan::Get(local, Nan::New("VERSION").ToLocalChecked());
  if (!ver.IsEmpty()) {
    return info.GetReturnValue().Set(ver.ToLocalChecked());
  }
}

NAN_METHOD(Method_printACP) // UTF-8 to MBCS (.ACP)
{
//  Nan::HandleScope scope; -- should be implicit in method calls
  if(info.Length() >= 1){
    String::Utf8Value s(info[0]);
    char *cs = *s;
    printf(UTF82MBCS(std::string(cs)).c_str());
  }
  return info.GetReturnValue().Set(true);
}

NAN_METHOD(Method_print) // through (as ASCII)
{
//  Nan::HandleScope scope; -- should be implicit in method calls
  if(info.Length() >= 1){
    String::Utf8Value s(info[0]);
    char *cs = *s;
    printf(cs); // printf("%s\n", cs);
  }
  return info.GetReturnValue().Set(true);
}

} // namespace node_win32ole

using namespace node_win32ole;

namespace {

NAN_MODULE_INIT(init)
{
  module_target.Reset(target);
  V8Variant::Init(target);
  Client::Init(target);
  Nan::ForceSet(target, Nan::New("VERSION").ToLocalChecked(),
    Nan::New("0.0.0 (will be set later)").ToLocalChecked(),
    static_cast<PropertyAttribute>(DontDelete));
  Nan::ForceSet(target, Nan::New("MODULEDIRNAME").ToLocalChecked(),
    Nan::New("/tmp").ToLocalChecked(),
    static_cast<PropertyAttribute>(DontDelete));
  Nan::ForceSet(target, Nan::New("SOURCE_TIMESTAMP").ToLocalChecked(),
    Nan::New(__FILE__ " " __DATE__ " " __TIME__).ToLocalChecked(),
    static_cast<PropertyAttribute>(ReadOnly | DontDelete));
  Nan::Export(target, "version", Method_version);
  Nan::Export(target, "printACP", Method_printACP);
  Nan::Export(target, "print", Method_print);
  Nan::Export(target, "gettimeofday", Method_gettimeofday);
  Nan::Export(target, "sleep", Method_sleep);
  Nan::Export(target, "force_gc_extension", Method_force_gc_extension);
  Nan::Export(target, "force_gc_internal", Method_force_gc_internal);
}

} // namespace

NODE_MODULE(node_win32ole, init)
