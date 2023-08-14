#include "xlwingsdll.h"

void ToVariant(const char* str, VARIANT* var)
{
	VariantClear(var);

	int sz = (int) strlen(str) + 1;
	OLECHAR* wide = new OLECHAR[sz];
	MultiByteToWideChar(CP_ACP, 0, str, sz * sizeof(OLECHAR), wide, sz);
	var->vt = VT_BSTR;
	var->bstrVal = SysAllocString(wide);
	delete[] wide;
}

void ToVariant(const std::string& str, VARIANT* var)
{
	ToVariant(str.c_str(), var);
}

void ToStdString(const wchar_t* ws, std::string& str)
{
	BOOL bUsedDefaultChar;
	int len = (int) wcslen(ws);
	char* narrow = new char[len+1];
	WideCharToMultiByte(CP_ACP, 0, ws, len, narrow, len+1, "?", &bUsedDefaultChar);
	narrow[len] = 0;
	str = narrow;
	delete narrow;
}

void ToStdString(BSTR bs, std::string& str)
{
	BOOL bUsedDefaultChar;
	int len = (int) SysStringLen(bs);
	AutoArrayDeleter<char> narrow(new char[len+1]);
	WideCharToMultiByte(CP_ACP, 0, bs, len, narrow.p, len+1, "?", &bUsedDefaultChar);
	narrow.p[len] = 0;
	str = narrow.p;
}

void ToBStr(const std::string& str, BSTR& bs)
{
	int sz = (int) str.length() + 1;
	OLECHAR* wide = new OLECHAR[sz];
	MultiByteToWideChar(CP_ACP, 0, str.c_str(), sz * sizeof(OLECHAR), wide, sz);
	bs = SysAllocString(wide);
	delete[] wide;
}

std::string GetLastErrorMessage()
{
	DWORD dwError = GetLastError();
	char* lpMsgBuf;
	
	if(0 == FormatMessageA(
		FORMAT_MESSAGE_ALLOCATE_BUFFER | 
		FORMAT_MESSAGE_FROM_SYSTEM |
		FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL,
		dwError,
		0,
		(LPSTR) &lpMsgBuf,
		0,
		NULL))
	{
		return "Could not get error message: FormatMessage failed.";
	}

	std::string ret = lpMsgBuf;
	LocalFree(lpMsgBuf);
	return ret;
}

const char* GetDLLPath()
{
	static bool initialized = false;
	static char path[MAX_PATH];

	if(!initialized)
	{
		if(0 == GetModuleFileNameA(hInstanceDLL, path, MAX_PATH))
			throw formatted_exception() << "GetModuleFileName failed.";

		initialized = true;
	}

	return path;
}


const char* GetDLLFolder()
{
	static bool initialized = false;
	static char folderPath[MAX_PATH];

	if(!initialized)
	{
		if(0 == GetModuleFileNameA(hInstanceDLL, folderPath, MAX_PATH))
			throw formatted_exception() << "GetModuleFileName failed.";

		int n = (int) strlen(folderPath) - 1;
		while(folderPath[n] != '\\' && n > 0)
			n--;

		if(n == 0)
			throw formatted_exception() << "Could deduce DLL folder, GetModuleFileName returned '" << folderPath << "'.";

		folderPath[n] = 0;

		initialized = true;
	}

	return folderPath;
}

void GetFullPathRelativeToDLLFolder(const std::string& path, std::string& out)
{
	char buffer[MAX_PATH];
	char curdir[MAX_PATH];
	if(0 == GetCurrentDirectoryA(MAX_PATH, curdir))
		throw formatted_exception() << "GetCurrentDirectory failed in GetFullPathRelativeToDLLFolder.";
	if(0 == SetCurrentDirectoryA(GetDLLFolder()))
		throw formatted_exception() << "SetCurrentDirectory (1st) failed in GetFullPathRelativeToDLLFolder.";
	if(0 == GetFullPathNameA(path.c_str(), MAX_PATH, buffer, NULL))
		throw formatted_exception() << "GetFullPathName failed in GetFullPathRelativeToDLLFolder.";
	if(0 == SetCurrentDirectoryA(curdir))
		throw formatted_exception() << "SetCurrentDirectory (2nd) failed in GetFullPathRelativeToDLLFolder.";
	out = buffer;
}

std::string GUIDToStdString(GUID& guid)
{
	char buffer[100];
	sprintf_s(buffer, 100, "{%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}",
          guid.Data1, guid.Data2, guid.Data3,
          guid.Data4[0], guid.Data4[1], guid.Data4[2],
          guid.Data4[3], guid.Data4[4], guid.Data4[5],
          guid.Data4[6], guid.Data4[7]);
    return buffer;
}

void ParseGUID(const char* s, GUID& guid)
{
    unsigned long p0;
    unsigned short p1, p2, p3, p4, p5, p6, p7, p8, p9, p10;

	int nConverted = sscanf_s(s, "{%8lX-%4hX-%4hX-%2hX%2hX-%2hX%2hX%2hX%2hX%2hX%2hX}", &p0, &p1, &p2, &p3, &p4, &p5, &p6, &p7, &p8, &p9, &p10);
	if(nConverted != 11)
		throw formatted_exception() << "Failed to parse GUID '" << s << "', it should be in the form {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx}.";

    guid.Data1 = p0;
    guid.Data2 = p1;
    guid.Data3 = p2;
    guid.Data4[0] = (unsigned char) p3;
    guid.Data4[1] = (unsigned char) p4;
    guid.Data4[2] = (unsigned char) p5;
    guid.Data4[3] = (unsigned char) p6;
    guid.Data4[4] = (unsigned char) p7;
    guid.Data4[5] = (unsigned char) p8;
    guid.Data4[6] = (unsigned char) p9;
    guid.Data4[7] = (unsigned char) p10;
}

void NewGUID(GUID& guid)
{
	HRESULT hr = CoCreateGuid(&guid);
	if(FAILED(hr))
		throw formatted_exception() << "CoCreateGuid failed.";
}

void GetLastWriteTime(const char* path, FILETIME* pFileTime)
{
	AutoCloseHandle hFile(CreateFileA(
		path,
		GENERIC_READ,
		FILE_SHARE_READ,
		NULL,
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL,
		NULL
		));

	if(hFile.handle == INVALID_HANDLE_VALUE)
		throw formatted_exception() << "Could not open file '" << path << "' to get last write time.";

	if(!GetFileTime(hFile.handle, NULL, NULL, pFileTime))
		throw formatted_exception() << "Could not get last write time for '" << path << "'.";
}
