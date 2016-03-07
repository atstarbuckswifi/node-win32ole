/*
  client.cc
*/

#include "client.h"
#include "v8dispatch.h"
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
    return Nan::ThrowError("Use the new operator to create new Client objects");
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
    return Nan::ThrowError("Can't create new Client object (null OLE32core)");
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
  if (info.Length() < 1 || !info[0]->IsString())
  {
    return Nan::ThrowTypeError("Argument 1 is not a String");
  }
  wchar_t *wcs;
  {
    String::Utf8Value u8s(info[0]); // must create here
    wcs = u8s2wcs(*u8s);
    if (!wcs) return Nan::ThrowError(NewOleException(GetLastError()));
  }
#ifdef DEBUG
  char *mbs = wcs2mbs(wcs);
  if (!mbs)
  {
    free(wcs);
    return Nan::ThrowError(NewOleException(GetLastError()));
  }
  fprintf(stderr, "ProgID: %s\n", mbs);
  free(mbs);
#endif
  CLSID clsid;
  HRESULT hr = CLSIDFromProgID(wcs, &clsid);
  free(wcs);
  if(FAILED(hr)) return Nan::ThrowError(NewOleException(hr));
#ifdef DEBUG
  fprintf(stderr, "clsid:"); // 00024500-0000-0000-c000-000000000046 (Excel) ok
  for(int i = 0; i < sizeof(CLSID); ++i)
    fprintf(stderr, " %02x", ((unsigned char *)&clsid)[i]);
  fprintf(stderr, "\n");
#endif
  Handle<Object> vApp = V8Dispatch::CreateNew(NULL);
  if (vApp.IsEmpty() || vApp->IsUndefined() || !vApp->IsObject())
  {
    return Nan::ThrowError("Unable to create return value placeholder");
  }
  V8Dispatch *v8d = V8Dispatch::Unwrap<V8Dispatch>(vApp);
  CHECK_V8(V8Dispatch, v8d);
  OCDispatch *app = &v8d->ocd;
  // When 'CoInitialize(NULL)' is not called first (and on the same instance),
  // next functions will return many errors.
  // (old style) GetActiveObject() returns 0x000036b7
  //   The requested lookup key was not found in any active activation context.
  // (OLE2) CoCreateInstance() returns 0x000003f0
  //   An attempt was made to reference a token that does not exist.
  // C -> C++ changes types (&clsid -> clsid, &IID_IDispatch -> IID_IDispatch)
  // options (CLSCTX_INPROC_SERVER CLSCTX_INPROC_HANDLER CLSCTX_LOCAL_SERVER)
  DWORD ctx = CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER;
  hr = CoCreateInstance(clsid, NULL, ctx, IID_IDispatch, (void **)&app->disp);
  if(FAILED(hr)){
    // Retry with WOW6432 bridge option.
    // This may not be a right way, but better.
    BDISPFUNCDAT("FAILED CoCreateInstance: %d: 0x%08x\n", 0, hr);
#if defined(_WIN64)
    ctx |= CLSCTX_ACTIVATE_32_BIT_SERVER; // 32bit COM server on 64bit OS
#else
    ctx |= CLSCTX_ACTIVATE_64_BIT_SERVER; // 64bit COM server on 32bit OS
#endif
    hr = CoCreateInstance(clsid, NULL, ctx, IID_IDispatch, (void **)&app->disp);
    if (FAILED(hr)) return Nan::ThrowError(NewOleException(hr));
  }
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
