// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the EXCELFUNCTIONSTEST_EXPORTS
// symbol defined on the command line. This symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// EXCELFUNCTIONSTEST_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef EXCELFUNCTIONSTEST_EXPORTS
#define EXCELFUNCTIONSTEST_API __declspec(dllexport)
#else
#define EXCELFUNCTIONSTEST_API __declspec(dllimport)
#endif

// This class is exported from the ExcelFunctionsTest.dll
class EXCELFUNCTIONSTEST_API CExcelFunctionsTest {
public:
	CExcelFunctionsTest(void);
	// TODO: add your methods here.
};

extern EXCELFUNCTIONSTEST_API int nExcelFunctionsTest;

EXCELFUNCTIONSTEST_API int fnExcelFunctionsTest(void);
