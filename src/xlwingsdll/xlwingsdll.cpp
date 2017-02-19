#include "xlwingsdll.h"

HINSTANCE hInstanceDLL;


// DLL entry point -- only stores the module handle for later use
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		hInstanceDLL = hinstDLL;
		return TRUE;

	case DLL_PROCESS_DETACH:
		Config::ClearConfigs();
		return TRUE;

	default:
		return TRUE;
	}
}


// change the dimensionality of a VBA array
HRESULT __stdcall XLPyDLLNDims(VARIANT* xlSource, int* xlDimension, bool *xlTranspose, VARIANT* xlDest)
{
	VariantClear(xlDest);

	try
	{
		// determine source dimensions
		int nSrcDims;
		int nSrcRows;
		int nSrcCols;
		SAFEARRAY* pSrcSA;
		if(0 == (xlSource->vt & VT_ARRAY))
		{
			nSrcDims = 0;
			nSrcRows = 1;
			nSrcCols = 1;
			pSrcSA = NULL;
		}
		else
		{
			if(xlSource->vt != (VT_VARIANT | VT_ARRAY) && xlSource->vt != (VT_VARIANT | VT_ARRAY | VT_BYREF))
				throw formatted_exception() << "Source array must be array of variants.";

			pSrcSA = (xlSource->vt & VT_BYREF) ? *xlSource->pparray : xlSource->parray;
			if(pSrcSA->cDims == 1)
			{
				nSrcDims = 1;
				nSrcRows = (int) pSrcSA->rgsabound->cElements;
				nSrcCols = 1;
			}
			else if(pSrcSA->cDims == 2)
			{
				nSrcDims = 2;
				nSrcRows = (int) pSrcSA->rgsabound[0].cElements;
				nSrcCols = (int) pSrcSA->rgsabound[1].cElements;
			}
			else
				throw formatted_exception() << "Source array must be either 1- or 2-dimensional.";
		}

		// determine dest dimension
		int nDestDims = *xlDimension;
		int nDestRows = nSrcRows;
		int nDestCols = nSrcCols;
		if(nDestDims == -1)
		{
			if(nSrcCols == 1 && nSrcRows == 1)
				nDestDims = 0;
			else if(nSrcCols == 1 || nSrcRows == 1)
				nDestDims = 1;
			else
				nDestDims = 2;
		}
		if(nDestDims == 1 && nSrcDims == 2)
		{
			if(nSrcRows != 1 && nSrcCols != 1)
				throw formatted_exception() << "When converting from 2- to 1-dimensional array, source must be (1 x n) or (n x 1).";
			if(nSrcCols != 1)
			{
				nDestRows = nSrcCols;
				nDestCols = nSrcRows;
			}
		}
		if(nDestDims == 0 && nSrcRows * nSrcCols != 1)
			throw formatted_exception() << "When converting array to scalar, source must contain only one element.";
		if(nDestDims == 2 && *xlTranspose)
		{
			int tmp = nDestRows;
			nDestRows = nDestCols;
			nDestCols = tmp;
		}

		// create destination safe array -- note that if nDestDims == 0 then AutoSafeArrayCreate does nothing and leaves pointer == NULL
		// for some reason, with 2-dimensional array bounds get swapped around by SafeArrayCreate
		SAFEARRAYBOUND bounds[2];
		bounds[0].lLbound = 1;
		bounds[0].cElements = (ULONG) nDestDims == 2 ? nDestCols : nDestRows;
		bounds[1].lLbound = 1;
		bounds[1].cElements = (ULONG) nDestDims == 2 ? nDestRows : nDestCols;
		AutoSafeArrayCreate asac(VT_VARIANT, nDestDims, bounds);

		// copy the data -- note that if NULL is passed to AutoSafeArrayAccessData then it does nothing
		{
			VARIANT* pSrcData = xlSource;
			AutoSafeArrayAccessData(pSrcSA, (void**) &pSrcData);

			VARIANT* pDestData = xlDest;
			AutoSafeArrayAccessData(asac.pSafeArray, (void**) &pDestData);

			for(int iDestRow=0; iDestRow<nDestRows; iDestRow++)
			{
				for(int iDestCol=0; iDestCol<nDestCols; iDestCol++)
				{
					// safe array data is column-major -- and treating col-major data as row-major is the same as transposing
					int destIdx = iDestRow * nDestCols + iDestCol;
					int srcIdx = *xlTranspose ? iDestRow + iDestCol * nDestRows : destIdx;

					VariantInit(&pDestData[destIdx]);
					VariantCopy(&pDestData[destIdx], &pSrcData[srcIdx]);
				}
			}
		}

		// if we created a safe array, return and release it
		if(asac.pSafeArray != NULL)
		{
			xlDest->vt = VT_VARIANT | VT_ARRAY;
			xlDest->parray = asac.pSafeArray;
			asac.pSafeArray = NULL;
		}

		return S_OK;
	}
	catch(const std::exception& e)
	{
		ToVariant(e.what(), xlDest);
		return E_FAIL;
	}
}


// entry point - returns existing interface if already available, otherwise tries to activate it
HRESULT __stdcall XLPyDLLActivate(VARIANT* xlResult, const char* xlConfigFileName, int xlActivationMode)
{
	try
	{
		VariantClear(xlResult);

		// set default config file
		std::string configFilename = xlConfigFileName;
		if(configFilename.empty())
			throw formatted_exception() << "No config file specified";
		Config* pConfig = Config::GetConfig(configFilename);

		// if interface object isn't already available try to create it
		switch(xlActivationMode)
		{
		case -1:
			{
				pConfig->KillRPCServer();
				return S_OK;
			}

		case 0:
			{
				if(pConfig->pInterface != NULL && pConfig->CheckRPCServer())
				{
					// pass it back to VBA
					xlResult->vt = VT_DISPATCH;
					xlResult->pdispVal = pConfig->pInterface;
					xlResult->pdispVal->AddRef();
					return S_OK;
				}
				else
				{
					xlResult->vt = VT_DISPATCH;
					xlResult->pdispVal = NULL;
					return S_OK;
				}
			}

		case 1:
			{
				if(pConfig->pInterface == NULL || !pConfig->CheckRPCServer())
					pConfig->ActivateRPCServer();

				// pass it back to VBA
				xlResult->vt = VT_DISPATCH;
				xlResult->pdispVal = pConfig->pInterface;
				xlResult->pdispVal->AddRef();

				return S_OK;
			}

		default:
			throw formatted_exception() << "Invalid value " << xlActivationMode << " for xlActivationMode, must be -1, 0 or 1";
		}
	}
	catch(const std::exception& e)
	{
		ToVariant(e.what(), xlResult);
		return E_FAIL;
	}
}


// special configuration-less entry point for used by xlwings
// just pass the command to launch the COM server, all other settings normally contained in the config file are default
// returns existing interface if already available, otherwise tries to activate it
HRESULT __stdcall XLPyDLLActivateAuto(VARIANT* xlResult, const char* xlCommand, int xlActivationMode)
{
	try
	{
		VariantClear(xlResult);

		// set default config file
		std::string command = xlCommand;
		if(command.empty())
			throw formatted_exception() << "No command line specified";
		Config* pConfig = Config::GetAutoConfig(command);

		// if interface object isn't already available try to create it

		switch (xlActivationMode)
		{
		case -1:
		{
			pConfig->KillRPCServer();
			return S_OK;
		}

		case 0:
		{
			if (pConfig->pInterface != NULL && pConfig->CheckRPCServer())
			{
				// pass it back to VBA
				xlResult->vt = VT_DISPATCH;
				xlResult->pdispVal = pConfig->pInterface;
				xlResult->pdispVal->AddRef();
				return S_OK;
			}
			else
			{
				xlResult->vt = VT_DISPATCH;
				xlResult->pdispVal = NULL;
				return S_OK;
			}
		}

		case 1:
		{
			if (pConfig->pInterface == NULL || !pConfig->CheckRPCServer())
				pConfig->ActivateRPCServer();

			// pass it back to VBA
			xlResult->vt = VT_DISPATCH;
			xlResult->pdispVal = pConfig->pInterface;
			xlResult->pdispVal->AddRef();

			return S_OK;
		}

		default:
			throw formatted_exception() << "Invalid value " << xlActivationMode << " for xlActivationMode, must be -1, 0 or 1";
		}
	}
	catch(const std::exception& e)
	{
		ToVariant(e.what(), xlResult);
		return E_FAIL;
	}
}


// returns a string identifying the DLL version
int __stdcall XLPyDLLVersion(BSTR* xlTag, double* xlVersion, BSTR* xlArchitecture)
{
	std::string tag("xlwingsdll");
	double version = 1.0;

#if _WIN64
	std::string arch("win64");
#else
	std::string arch("win32");
#endif

	if(xlTag != NULL)
	{
		SysFreeString(*xlTag);
		*xlTag = SysAllocStringByteLen(tag.c_str(), (UINT) tag.size());
	}

	if(xlVersion != NULL)
		*xlVersion = version;

	if(xlArchitecture != NULL)
	{
		SysFreeString(*xlArchitecture);
		*xlArchitecture = SysAllocStringByteLen(arch.c_str(), (UINT) arch.size());
	}

	return 1;
}
