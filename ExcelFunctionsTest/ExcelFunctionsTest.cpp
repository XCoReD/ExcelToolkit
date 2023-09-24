// ExcelFunctionsTest.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "ExcelFunctionsTest.h"


// This is an example of an exported variable
EXCELFUNCTIONSTEST_API int nExcelFunctionsTest=0;

// This is an example of an exported function.
EXCELFUNCTIONSTEST_API int fnExcelFunctionsTest(void)
{
    return 42;
}

// This is the constructor of a class that has been exported.
// see ExcelFunctionsTest.h for the class definition
CExcelFunctionsTest::CExcelFunctionsTest()
{
    return;
}
