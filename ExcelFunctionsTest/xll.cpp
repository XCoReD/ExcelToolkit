# include <windows.h>
# include <iostream>

//https://stackoverflow.com/questions/1940747/calling-excel-dll-xll-functions-from-c-sharp/2865023#2865023

// pointer to function taking 2 XLOPERS
typedef XLOPER (__cdecl *xl2args) (XLOPER, XLOPER);

void test(LPCTSTR library) 
{
	/// get the XLL address
	HINSTANCE h = LoadLibrary(library);
	if (h != NULL) 
	{
		xl2args myfunc;
		/// get my xll-dll.xll function address
		myfunc = (xl2args)GetProcAddress(h, "f0");
		if (!myfunc) 
		{ // handle the error
			FreeLibrary(h);
		}
		else 
		{ /// build some XLOPERS, call the remote function
			XLOPER a, b, *c;
			a.xltype = 1; a.val.num = 1.;
			b.xltype = 1; b.val.num = 2.;
			c = (*myfunc)(&a, &b);
			std::cout << " call of xll " << c->val.num << std::endl;
		}
		FreeLibrary(h);
	}
}
