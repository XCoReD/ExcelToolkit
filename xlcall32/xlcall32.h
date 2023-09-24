// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the XLCALL32_EXPORTS
// symbol defined on the command line. This symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// XLCALL32_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef XLCALL32_EXPORTS
#define XLCALL32_API __declspec(dllexport)
#else
#define XLCALL32_API __declspec(dllimport)
#endif

// This class is exported from the xlcall32.dll
class XLCALL32_API Cxlcall32 {
public:
	Cxlcall32(void);
	// TODO: add your methods here.
};

extern XLCALL32_API int nxlcall32;

XLCALL32_API int fnxlcall32(void);
