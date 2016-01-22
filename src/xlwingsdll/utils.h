void ToVariant(const char* str, VARIANT* var);
void ToVariant(const std::string& str, VARIANT* var);
void ToStdString(const wchar_t* ws, std::string& str);
void ToStdString(BSTR bs, std::string& str);
void ToBStr(const std::string& str, BSTR& bs);

class formatted_exception : public std::exception
{
protected:
	std::string s;

public:
	template<typename T>
	formatted_exception& operator<< (const T& value)
	{
		std::ostringstream oss;
		oss << value;
		s += oss.str();
		return *this;
	}

	formatted_exception& operator<< (const wchar_t* value)
	{
		std::string str;
		ToStdString(value, str);
		s += str;
		return *this;
	}

	virtual const char* what() const
	{
		return s.c_str();
	}
};

const char* GetDLLPath();
const char* GetDLLFolder();
void GetFullPathRelativeToDLLFolder(const std::string& path, std::string& out);

std::string GUIDToStdString(GUID& guid);
void ParseGUID(const char* str, GUID& guid);
void NewGUID(GUID& guid);

void GetLastWriteTime(const char* path, FILETIME* pFileTime);

std::string GetLastErrorMessage();

static inline std::string strlower(std::string& s)
{
	for (size_t k = 0; k < s.length(); k++)
		s[k] = std::tolower(s[k], std::locale());
	return s;
}

static inline std::string strupper(std::string& s)
{
	for (size_t k = 0; k < s.length(); k++)
		s[k] = std::toupper(s[k], std::locale());
	return s;
}

static inline std::string strtrim(std::string &s)
{
	s.erase(s.begin(), std::find_if(s.begin(), s.end(), std::not1(std::ptr_fun<int, int>(std::isspace))));
	s.erase(std::find_if(s.rbegin(), s.rend(), std::not1(std::ptr_fun<int, int>(std::isspace))).base(), s.end());
	return s;
}

static inline void strsplit(const std::string& s, const std::string& sep, std::vector<std::string>& out, bool trim=true)
{
	std::string ss = s;
	size_t pos;
	while(std::string::npos != (pos = ss.find(sep)))
	{
		out.push_back(trim ? strtrim(ss.substr(0, pos)) : ss.substr(0, pos));
		ss = ss.substr(pos + sep.length());
	}
	out.push_back(trim ? strtrim(ss) : ss);
}

template<class T>
class AutoArrayDeleter
{
public:
	T* p;
	AutoArrayDeleter(T* p)
	{
		this->p = p;
	}
	~AutoArrayDeleter()
	{
		delete[] p;
	}
};

class AutoSafeArrayAccessData
{
	SAFEARRAY* _pSA;

public:
	AutoSafeArrayAccessData(SAFEARRAY* pSA, void** ppData)
	{
		_pSA = NULL;
		if(pSA != NULL)
		{
			if(FAILED(SafeArrayAccessData(pSA, ppData)))
				throw formatted_exception() << "Could not access safe array data.";
		}
		_pSA = pSA;
	}

	~AutoSafeArrayAccessData()
	{
		if(_pSA != NULL && FAILED(SafeArrayUnaccessData(_pSA)))
			throw formatted_exception() << "Could not unaccess safe array data.";
	}
};

class AutoSafeArrayCreate
{
public:
	SAFEARRAY* pSafeArray;

	AutoSafeArrayCreate(VARTYPE vt, UINT cDims, SAFEARRAYBOUND* rgsabound)
		: pSafeArray(NULL)
	{
		if(cDims > 0)
		{
			pSafeArray = SafeArrayCreate(vt, cDims, rgsabound);
			if(pSafeArray == NULL)
				throw formatted_exception() << "Could not create safe array.";
		}
	}

	~AutoSafeArrayCreate()
	{
		if(pSafeArray != NULL && FAILED(SafeArrayDestroy(pSafeArray)))
			throw formatted_exception() << "Could not destroy safe array.";
	}
};

class AutoCloseHandle
{
public:
	HANDLE handle;

	AutoCloseHandle(const HANDLE& handle)
	{
		this->handle = handle;
	}

	~AutoCloseHandle()
	{
		if(this->handle != NULL && this->handle != INVALID_HANDLE_VALUE)
			CloseHandle(this->handle);
	}

	operator HANDLE ()
	{
		return handle;
	}
};
