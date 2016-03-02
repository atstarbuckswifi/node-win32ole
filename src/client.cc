/*
  client.cc
*/

#include "client.h"
#include "v8variant.h"

using namespace v8;
using namespace ole32core;

namespace node_win32ole {

Nan::Persistent<FunctionTemplate> Client::clazz;

void Client::Init(Nan::ADDON_REGISTER_FUNCTION_ARGS_TYPE target)
{
  Nan::HandleScope scope;
  Local<FunctionTemplate> t = Nan::New<FunctionTemplate>(New);
  t->InstanceTemplate()->SetInternalFieldCount(1);
  t->SetClassName(Nan::New("Client").ToLocalChecked());
//  Nan::SetPrototypeMethod(t, "New", New);
  Nan::SetPrototypeMethod(t, "Dispatch", Dispatch);
  Nan::SetPrototypeMethod(t, "Finalize", Finalize);
  Nan::Set(target, Nan::New("Client").ToLocalChecked(), Nan::GetFunction(t).ToLocalChecked());
  clazz.Reset(t);
}

NAN_METHOD(Client::New)
{
  DISPFUNCIN();
  if(!info.IsConstructCall())
    return Nan::ThrowTypeError("Use the new operator to create new Client objects");
  std::string cstr_locale(".ACP"); // default
  if(info.Length() >= 1){
    if(!info[0]->IsString())
      return Nan::ThrowTypeError("Argument 1 is not a String");
    String::Utf8Value u8s_locale(info[0]);
    cstr_locale = std::string(*u8s_locale);
  }
  Local<Object> thisObject = info.This();
  Client *cl = new Client(); // must catch exception
  if (!cl)
    return Nan::ThrowTypeError("Can't create new Client object (null OLE32core)");
  cl->Wrap(thisObject); // InternalField[0]
  HRESULT cnresult = cl->oc.connect(cstr_locale);
  if (FAILED(cnresult))
  {
    return Nan::ThrowError(NewOleException(cnresult));
  }
  DISPFUNCOUT();
  return info.GetReturnValue().Set(thisObject);
}

NAN_METHOD(Client::Dispatch)
{
  DISPFUNCIN();
  BEVERIFY(done, info.Length() >= 1);
  BEVERIFY(done, info[0]->IsString());
  if (false)
  {
done:
    DISPFUNCOUT();
    return Nan::ThrowTypeError("Dispatch failed");
  }
  wchar_t *wcs;
  {
    String::Utf8Value u8s(info[0]); // must create here
    wcs = u8s2wcs(*u8s);
  }
  BEVERIFY(done, wcs);
#ifdef DEBUG
  char *mbs = wcs2mbs(wcs);
  if(!mbs) free(wcs);
  BEVERIFY(done, mbs);
  fprintf(stderr, "ProgID: %s\n", mbs);
  free(mbs);
#endif
  CLSID clsid;
  HRESULT hr = CLSIDFromProgID(wcs, &clsid);
  free(wcs);
  BEVERIFY(done, !FAILED(hr));
#ifdef DEBUG
  fprintf(stderr, "clsid:"); // 00024500-0000-0000-c000-000000000046 (Excel) ok
  for(int i = 0; i < sizeof(CLSID); ++i)
    fprintf(stderr, " %02x", ((unsigned char *)&clsid)[i]);
  fprintf(stderr, "\n");
#endif
  Handle<Object> vApp = V8Variant::CreateUndefined();
  BEVERIFY(done, !vApp.IsEmpty());
  BEVERIFY(done, !vApp->IsUndefined());
  BEVERIFY(done, vApp->IsObject());
  V8Variant *v8v = V8Variant::Unwrap<V8Variant>(vApp);
  CHECK_V8V(v8v);
  OCVariant *app = &v8v->ocv;
  app->v.vt = VT_DISPATCH;
  // When 'CoInitialize(NULL)' is not called first (and on the same instance),
  // next functions will return many errors.
  // (old style) GetActiveObject() returns 0x000036b7
  //   The requested lookup key was not found in any active activation context.
  // (OLE2) CoCreateInstance() returns 0x000003f0
  //   An attempt was made to reference a token that does not exist.
  REFIID riid = IID_IDispatch; // can't connect to Excel etc with IID_IUnknown
  // C -> C++ changes types (&clsid -> clsid, &IID_IDispatch -> IID_IDispatch)
  // options (CLSCTX_INPROC_SERVER CLSCTX_INPROC_HANDLER CLSCTX_LOCAL_SERVER)
  DWORD ctx = CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER;
  hr = CoCreateInstance(clsid, NULL, ctx, riid, (void **)&app->v.pdispVal);
  if(FAILED(hr)){
    // Retry with WOW6432 bridge option.
    // This may not be a right way, but better.
    BDISPFUNCDAT("FAILED CoCreateInstance: %d: 0x%08x\n", 0, hr);
#if defined(_WIN64)
    ctx |= CLSCTX_ACTIVATE_32_BIT_SERVER; // 32bit COM server on 64bit OS
#else
    ctx |= CLSCTX_ACTIVATE_64_BIT_SERVER; // 64bit COM server on 32bit OS
#endif
    hr = CoCreateInstance(clsid, NULL, ctx, riid, (void **)&app->v.pdispVal);
  }
  if (FAILED(hr)) return Nan::ThrowError(NewOleException(hr));
  DISPFUNCOUT();
  return info.GetReturnValue().Set(vApp);
}

NAN_METHOD(Client::Finalize)
{
  DISPFUNCIN();
#if(0)
  std::cerr << __FUNCTION__ << " Finalizer is called\a" << std::endl;
  std::cerr.flush();
#endif
  Client *cl = Client::Unwrap<Client>(info.This());
  if (cl) cl->Finalize();
  DISPFUNCOUT();
  return info.GetReturnValue().Set(info.This());
}

void Client::Finalize()
{
  if (!finalized)
  {
    oc.disconnect();
    finalized = true;
  }
}

} // namespace node_win32ole
