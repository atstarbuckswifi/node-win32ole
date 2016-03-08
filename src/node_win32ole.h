#ifndef __NODE_WIN32OLE_H__
#define __NODE_WIN32OLE_H__

#include <node_buffer.h>
#include <node.h>
#include <nan.h>
#include <v8.h>

using namespace v8;

namespace node_win32ole {

#define CHECK_V8(cls,v8) do{ \
    if(!(v8)) \
      return Nan::ThrowError( __FUNCTION__ " can't access to " #cls); \
  }while(0)
#define CHECK_V8_UNDEFINED(cls,v8) do{ \
    if(!(v8)) \
    { \
      Nan::ThrowError( __FUNCTION__ " can't access to " #cls); \
      return Nan::Undefined(); \
    } \
  }while(0)

#if(DEBUG)
#define OLETRACEIN() BDISPFUNCIN()
#define OLETRACEARG(v) do{ \
    std::cerr << (v->IsObject() ? "OBJECT" : *String::Utf8Value(v)) << ","; \
  }while(0)
#define OLETRACEPREARGV(sargs) Handle<Value> argv[] = { sargs }; \
  int argc = sizeof(argv) / sizeof(argv[0])
#define OLETRACEARGV() do{ \
    for(int i = 0; i < argc; ++i) OLETRACEARG(argv[i]); \
  }while(0)
#define OLETRACEARGS() do{ \
    for(int i = 0; i < info.Length(); ++i) OLETRACEARG(info[i]); \
  }while(0)
#define OLETRACEFLUSH() do{ std::cerr<<std::endl; std::cerr.flush(); }while(0)
#define OLETRACEOUT() BDISPFUNCOUT()
#else
#define OLETRACEIN()
#define OLETRACEARG(v)
#define OLETRACEPREARGV(sargs)
#define OLETRACEARGV()
#define OLETRACEARGS()
#define OLETRACEFLUSH()
#define OLETRACEOUT()
#endif

#define GET_PROP(obj, prop) Nan::Get((obj), Nan::New<String>(prop).ToLocalChecked())

#define ARRAY_AT(a, i) ((a)->Get(Nan::GetCurrentContext(), Nan::New<String>(to_s(i).c_str()).ToLocalChecked()).ToLocalChecked())
#define ARRAY_SET(a, i, v) Nan::Set((a), Nan::New<String>(to_s(i).c_str()).ToLocalChecked(), (v))

#define INSTANCE_CALL(obj, method, argc, argv) Handle<Function>::Cast( \
  GET_PROP((obj), (method)).ToLocalChecked())->Call((obj), (argc), (argv))

extern Nan::Persistent<Object> module_target;

NAN_METHOD(Method_gettimeofday);
NAN_METHOD(Method_sleep); // ms, bool: msg, bool: \n
NAN_METHOD(Method_force_gc_extension); // v8/gc : gc()
NAN_METHOD(Method_force_gc_internal);

} // namespace node_win32ole

#endif // __NODE_WIN32OLE_H__
