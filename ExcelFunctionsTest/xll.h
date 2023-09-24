#pragma once

# include <windows.h>

typedef void* LPXLOPER;

extern "C" void __declspec(dllexport) XLCallVer() {}

extern "C" int __declspec(dllexport) Excel4(int xlfn, LPXLOPER operRes, int count, ...) { return 0; }

extern "C" int __declspec(dllexport) Excel4v(int xlfn, LPXLOPER operRes, int count, LPXLOPER far opers[]) { return 0; }