/*
  force_gc_internal.cc
*/

#include "node_win32ole.h"
#include <iostream>
#include <node.h>
#include <nan.h>

using namespace v8;

namespace node_win32ole {

NAN_METHOD(Method_force_gc_internal)
{
  std::cerr << "-in: " __FUNCTION__ << std::endl;
  if(info.Length() < 1)
    return Nan::ThrowTypeError("this function takes at least 1 argument(s)");
  if(!info[0]->IsInt32())
    return Nan::ThrowTypeError("type of argument 1 must be Int32");
  int flags = (int)info[0]->Int32Value();
  while (!Nan::IdleNotification(100)){}
  std::cerr << "-out: " __FUNCTION__ << std::endl;
  return info.GetReturnValue().Set(true);
}

} // namespace node_win32ole
