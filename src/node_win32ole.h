#ifndef __NODE_WIN32OLE_H__
#define __NODE_WIN32OLE_H__

#include <node_buffer.h>
#include <node.h>
#include <nan.h>
#include <v8.h>

using namespace v8;

namespace node_win32ole {

#define CHECK_OCV(ocv) do{ \
    if(!(ocv)) \
      return Nan::ThrowError( __FUNCTION__" can't access to V8Variant (null OCVariant)"); \
  }while(0)
#define CHECK_OCV_UNDEFINED(ocv) do{ \
    if(!(ocv)) \
      Nan::ThrowError( __FUNCTION__" can't access to V8Variant (null OCVariant)"); \
      return Nan::Undefined(); \
  }while(0)

#define CHECK_V8V(v8v) do{ \
    if(!(v8v)) \
      return Nan::ThrowError( __FUNCTION__" can't access to V8Variant"); \
  }while(0)
#define CHECK_V8V_UNDEFINED(v8v) do{ \
    if(!(v8v)) \
      Nan::ThrowError( __FUNCTION__" can't access to V8Variant"); \
      return Nan::Undefined(); \
  }while(0)

#if(DEBUG)
#define OLETRACEIN() BDISPFUNCIN()
#define OLETRACEVT(th) do{ \
    V8Variant *v8v = V8Variant::Unwrap<V8Variant>(th); \
    if(!v8v){ std::cerr << "*** V8Variant is NULL ***"; std::cerr.flush(); } \
    CHECK_V8V(v8v); \
    OCVariant *ocv = &v8v->ocv; \
    std::cerr << "0x" << std::setw(8) << std::left << std::hex << ocv << ":"; \
    std::cerr << "vt=" << ocv->v.vt << ":"; \
    std::cerr.flush(); \
  }while(0)
#define OLETRACEVT_UNDEFINED(th) do{ \
    V8Variant *v8v = V8Variant::Unwrap<V8Variant>(th); \
    if(!v8v){ std::cerr << "*** V8Variant is NULL ***"; std::cerr.flush(); } \
    CHECK_V8V_UNDEFINED(v8v); \
    OCVariant *ocv = &v8v->ocv; \
    std::cerr << "0x" << std::setw(8) << std::left << std::hex << ocv << ":"; \
    std::cerr << "vt=" << ocv->v.vt << ":"; \
    std::cerr.flush(); \
  }while(0)
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
#define OLE_PROCESS_CARRY_OVER(th) do{ \
    V8Variant *v8v = node::ObjectWrap::Unwrap<V8Variant>(th); \
    if(v8v->property_carryover.empty()) break; \
    Handle<Value> r = V8Variant::OLEFlushCarryOver(th); \
    if(r->IsUndefined()) return; \
    if(!r->IsObject()){ \
      std::cerr << "** CarryOver primitive ** " << __FUNCTION__ << std::endl; \
      std::cerr.flush(); \
      return info.GetReturnValue().Set(r); \
    } \
    th = Nan::To<Object>(r).ToLocalChecked(); \
  }while(0)
#else
#define OLETRACEIN()
#define OLETRACEVT(th)
#define OLETRACEVT_UNDEFINED(th)
#define OLETRACEARG(v)
#define OLETRACEPREARGV(sargs)
#define OLETRACEARGV()
#define OLETRACEARGS()
#define OLETRACEFLUSH()
#define OLETRACEOUT()
#define OLE_PROCESS_CARRY_OVER(th) do{ \
    V8Variant *v8v = node::ObjectWrap::Unwrap<V8Variant>(th); \
    if(v8v->property_carryover.empty()) break; \
    Handle<Value> r = V8Variant::OLEFlushCarryOver(th); \
    if(r->IsUndefined()) return; \
    if(!r->IsObject()) return info.GetReturnValue().Set(r); \
    th = Nan::To<Object>(r).ToLocalChecked(); \
  }while(0)
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
