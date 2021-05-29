// pch.cpp: source file corresponding to the pre-compiled header

#include "pch.h"

// A CLSID is required for registration of the extension in the registry.
// This extension's CLSID is: { e0b60a9b-9f69-4ca6-ac9b-edbb502a9fce }
const CLSID CLSID_OperaFSShellExtension =
{ 0xdeadbeef, 0xdead, 0xbeef, { 0xde, 0xad, 0xbe, 0xef, 0xde, 0xad, 0xbe, 0xef } };
// When you are using pre-compiled headers, this source file is necessary for compilation to succeed.
