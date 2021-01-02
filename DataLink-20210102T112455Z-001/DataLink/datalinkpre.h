#ifndef __DataLink_PRE__

#define __DataLink_PRE__

#include <Windows.h>
#include "DataLinkGeneral.h"


#ifdef __cplusplus
	#define DL_EXPORT	extern "C" __declspec(dllexport)
#else
	#define DL_EXPORT	__declspec(dllexport)
#endif

extern HINSTANCE hDLinstance;

#define DL_REG_KEY	"Software\\Diebold\\Agilis Power\\DataLink"

#define ENCAPSULATE_ARRAY(VarName, VarType, Size, Default)			\
																	\
	struct ST_##VarName												\
	{																\
		VarType	VarName[Size];										\
	} st##VarName##default = { Default };


#define DECLARE_SHARED_MEM_ARRAY(VarName, Size, VarDefault)					\
	ENCAPSULATE_ARRAY	(##VarName, char, ##Size, ##VarDefault)				\
	DECLARE_VARIABLE	(##VarName, ST_##VarName, st##VarName##default)		\
	STORAGE_SHARED_MEM	(##VarName, ST_##VarName)							\
	NEEDS_FORMATTING	(##VarName, ST_##VarName)							\
	DEFAULT_FORMATTING	(##VarName, char)

#define DECLARE_SHARED_MEM_INT(VarName, Default)							\
	DECLARE_VARIABLE		(##VarName, int, ##Default)						\
	STORAGE_SHARED_MEM		(##VarName, int)								\
	NEEDS_FORMATTING		(##VarName, int)								\
	DEFAULT_FORMATTING		(##VarName, int)

#define DECLARE_WAITABLE_SHARED_MEM_ARRAY(VarName, Size, VarDefault)		\
	ENCAPSULATE_ARRAY	(##VarName, char, ##Size, ##VarDefault)				\
	DECLARE_WAITABLE_VARIABLE(##VarName, ST_##VarName, st##VarName##default)\
	STORAGE_SHARED_MEM	(##VarName, ST_##VarName)							\
	NEEDS_FORMATTING	(##VarName, ST_##VarName)							\
	DEFAULT_FORMATTING	(##VarName, char)

#define DECLARE_WAITABLE_SHARED_MEM_INT(VarName, VarDefault)				\
	DECLARE_WAITABLE_VARIABLE(##VarName, int, ##VarDefault)					\
	STORAGE_SHARED_MEM	(##VarName, int)									\
	NEEDS_FORMATTING	(##VarName, int)									\
	DEFAULT_FORMATTING	(##VarName, int)

#define DECLARE_REGISTRY_ARRAY(VarName, Size, VarDefault)					\
	ENCAPSULATE_ARRAY	(##VarName, char, ##Size, ##VarDefault)				\
	DECLARE_VARIABLE	(##VarName, ST_##VarName, st##VarName##default)		\
	STORAGE_REGISTRY	(##VarName, ST_##VarName)							\
	NEEDS_FORMATTING	(##VarName, ST_##VarName)							\
	DEFAULT_FORMATTING	(##VarName, char)

#define DECLARE_WAITABLE_REGISTRY_ARRAY(VarName, Size, VarDefault)			\
	ENCAPSULATE_ARRAY	(##VarName, char, ##Size, ##VarDefault)				\
	DECLARE_WAITABLE_VARIABLE(##VarName, ST_##VarName, st##VarName##default)\
	STORAGE_REGISTRY	(##VarName, ST_##VarName)							\
	NEEDS_FORMATTING	(##VarName, ST_##VarName)							\
	DEFAULT_FORMATTING	(##VarName, char)

#define DECLARE_REGISTRY_DOUBLE(VarName, VarDefault)						\
	DECLARE_VARIABLE	(##VarName, double, VarDefault)						\
	STORAGE_REGISTRY	(##VarName, double)

#define DECLARE_REGISTRY_INTEGER(VarName, Default)							\
	DECLARE_VARIABLE		(##VarName, int, ##Default)						\
	STORAGE_REGISTRY		(##VarName, int)								\
	NEEDS_FORMATTING		(##VarName, int)								\
	DEFAULT_FORMATTING		(##VarName, int)

#define DECLARE_VARIABLE(VarName, VarType, VarDefault)						\
																			\
	extern	short	Serialize##VarName(LPVOID, BOOL);						\
																			\
	DL_EXPORT	short	DLGet##VarName(LPVOID Destination)					\
	{																		\
		return Serialize##VarName(Destination, TRUE);						\
	}																		\
																			\
	DL_EXPORT	short	DLSet##VarName(LPVOID Value)						\
	{																		\
		return Serialize##VarName(Value, FALSE);							\
	}																		\
																			\
	DL_EXPORT	int		DLGetSize##VarName()								\
	{																		\
		return sizeof(VarType);												\
	}																		\
																			\
	DL_EXPORT	short	DLReset##VarName()									\
	{																		\
		VarType Var=VarDefault;												\
		return Serialize##VarName((LPVOID)&Var, FALSE);						\
	}

#define DECLARE_WAITABLE_VARIABLE(VarName, VarType, VarDefault)				\
	extern	short	Serialize##VarName(LPVOID, BOOL);						\
																			\
	DL_EXPORT	short	DLGet##VarName(LPVOID Destination)					\
	{																		\
		return Serialize##VarName(Destination, TRUE);						\
	}																		\
																			\
	DL_EXPORT	short	DLSet##VarName(LPVOID Value)						\
	{																		\
		short nReply=Serialize##VarName(Value, FALSE);						\
		HANDLE hEvent = CreateEvent( NULL, TRUE, FALSE, "AgilisDLWaitEvt" #VarName);	\
		if ( NULL != hEvent )												\
		{																	\
			PulseEvent( hEvent );											\
			CloseHandle( hEvent );											\
		}																	\
		return nReply;														\
	}																		\
																			\
	DL_EXPORT	int		DLGetSize##VarName()								\
	{																		\
		return sizeof(VarType);												\
	}																		\
																			\
	DL_EXPORT	short	DLReset##VarName()									\
	{																		\
		VarType Var=VarDefault;												\
		return Serialize##VarName((LPVOID)&Var, FALSE);						\
	}																		\
																			\
	DL_EXPORT	DWORD	DLWait##VarName(int nTimeout)						\
	{																		\
		DWORD dwReply=DL_WAIT_FAILED;										\
		HANDLE hEvent = CreateEvent( NULL, TRUE, FALSE, "AgilisDLWaitEvt" #VarName);	\
		if ( NULL != hEvent )												\
		{																	\
			dwReply = WaitForSingleObject( hEvent, nTimeout );				\
			CloseHandle( hEvent );											\
		}																	\
		return dwReply;														\
	}																		\



#define STORAGE_SHARED_MEM(VarName, VarType)								\
																			\
	VarType	VarName={0};													\
																			\
	short	Serialize##VarName(LPVOID lpBuffer, BOOL bRead)					\
	{																		\
		if(bRead)															\
		{																	\
			/* memcpy something */											\
			*(VarType *)lpBuffer=VarName;									\
		}																	\
		else																\
		{																	\
			/* memcpy something */											\
			VarName=*(VarType *)lpBuffer;									\
		}																	\
		return (DL_SUCCESS);												\
	}


#define STORAGE_REGISTRY(VarName, VarType)											\
																					\
	short	Serialize##VarName(LPVOID lpBuffer, BOOL bRead)							\
	{																				\
		HKEY hKey;																	\
		LONG Result;																\
																					\
		Result=RegOpenKey(HKEY_LOCAL_MACHINE, DL_REG_KEY, &hKey);					\
		if(ERROR_SUCCESS!=Result)													\
		{																			\
			return (DL_UNDEFINED);													\
		}																			\
																					\
		if(bRead)																	\
		{																			\
			DWORD dwSize=sizeof(VarType);											\
																					\
			Result=RegQueryValueEx(hKey, #VarName, NULL, NULL,						\
									(unsigned char *)lpBuffer, &dwSize);			\
			RegCloseKey(hKey);														\
			if(ERROR_SUCCESS!=Result)												\
			{																		\
				return (DL_UNDEFINED);												\
			}																		\
			else																	\
			{																		\
				return (DL_SUCCESS);												\
			}																		\
		}																			\
		else																		\
		{																			\
			Result=RegSetValueEx(hKey, #VarName, NULL, REG_BINARY,					\
								 (unsigned char *)lpBuffer, sizeof(VarType));		\
			RegCloseKey(hKey);														\
			if(ERROR_SUCCESS!=Result)												\
			{																		\
				return (DL_UNDEFINED);												\
			}																		\
			else																	\
			{																		\
				return (DL_SUCCESS);												\
			}																		\
		}																			\
	}


#define NEEDS_FORMATTING(VarName, VarType)											\
																					\
	extern	BOOL	IsInternalFormat##VarName(char *CheckType);/*PEC*/				\
  	extern	short	Format##VarName(LPVOID Internal, LPVOID External,				\
									LPSTR Format, int nDirectionn, int size);/*PEC*/\
																					\
	DL_EXPORT short	DLFGet##VarName(LPVOID Destination, LPSTR Format)			\
	{																				\
/* PEC */																			\
		if (IsInternalFormat##VarName(Format) )										\
		{																			\
			return DLGet##VarName(Destination);										\
		}																			\
																					\
		VarType *Var = new VarType;													\
		short Reply;																\
																					\
		DLGet##VarName(Var);														\
/* PEC */																			\
		Reply=Format##VarName((LPVOID)Var, Destination, Format,						\
								FORMAT_INTERNAL_TO_SPECIFIED, sizeof(VarType) );	\
		delete Var;																	\
		return Reply;																\
	}																				\
																					\
	DL_EXPORT	short	DLFSet##VarName(LPVOID Value, LPSTR Format)					\
	{																				\
/* PEC */																			\
		if (IsInternalFormat##VarName(Format) )										\
		{																			\
			return DLSet##VarName(Value);											\
		}																			\
																					\
		VarType *Var = new VarType;													\
		short Reply;																\
																					\
/* PEC */																			\
		Reply=Format##VarName((LPVOID)Var, Value, Format,							\
								FORMAT_SPECIFIED_TO_INTERNAL, sizeof(VarType) );	\
		if (DL_SUCCESS != Reply)													\
		{																			\
			delete Var;																\
			return Reply;															\
		}																			\
		Reply=DLSet##VarName(Var);													\
																					\
		delete Var;																	\
		return Reply;																\
	}


#define DEFAULT_FORMATTING(VarName, VarType)									\
																				\
	extern	short	 Format_##VarType(LPVOID, LPVOID, LPSTR, int, int);			\
																				\
	short	Format##VarName(LPVOID Internal, LPVOID External,					\
										LPSTR Format, int nDirection, int size)	\
	{																			\
		return Format_##VarType(Internal, External, Format, nDirection, size);	\
	}																			\
/*PEC*/																			\
	BOOL IsInternalFormat##VarName(char *CheckType)								\
	{																			\
		return ( ! _strcmpi (#VarType, CheckType) );					\
	}


#define CUSTOM_FORMATTING(VarName, VarType, InternalToSpecified, SpecifiedToInternal) \
																					\
	short	Format##VarName(LPVOID Internal, LPVOID External,						\
							LPSTR Format, int nDirection, int size)					\
	{																				\
		switch(nDirection)															\
		{																			\
			case FORMAT_INTERNAL_TO_SPECIFIED:										\
				return InternalToSpecified(Internal, External, Format, , int size);	\
			break;																	\
																					\
			case FORMAT_SPECIFIED_TO_INTERNAL:										\
/*PEC*/			return SpecifiedToInternal(External, Internal, Format, int size);	\
			break;																	\
																					\
			default:																\
				return (DL_UNDEFINED);												\
			break;																	\
		}																			\
	}


#pragma data_seg(".DLSM")
#undef	__DataLink_POST__


#endif  // __DataLink_PRE__
