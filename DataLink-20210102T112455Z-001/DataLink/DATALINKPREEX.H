#ifndef __DataLink_PRE_EX__

#define __DataLink_PRE_EX__

#include <Windows.h>
#include "DataLinkGeneral.h"


#ifdef __cplusplus
	#define DL_EXPORT	extern "C" __declspec(dllexport)
#else
	#define DL_EXPORT	__declspec(dllexport)
#endif

extern HINSTANCE hDLinstance;


//NOTE:
//All defines in this .h create variable that can't use the format facility
//To use it you mast modify this macros and add to the end (for example):
//
//	NEEDS_FORMATTING	(##VarName, ST_##VarName)
//	DEFAULT_FORMATTING	(##VarName, char)


#define DECLARE_REGISTRY_ARRAY(VarName, Size, VarDefault)					\
	ENCAPSULATE_ARRAY	(##VarName, char, ##Size, ##VarDefault)				\
	DECLARE_VARIABLE	(##VarName, ST_##VarName, st##VarName##default)		\
	STORAGE_REGISTRY	(##VarName, ST_##VarName)							\
//This macro is used to create a String in the registry

#define DECLARE_SHARED_MEM_VAR(VarName, VarType, VarDefault)				\
	DECLARE_VARIABLE	(##VarName, ##VarType, VarDefault)					\
	STORAGE_SHARED_MEM	(##VarName, ##VarType)								\
//This macro is used to create a variable not String in shared memory

#define DECLARE_REGISTRY_VAR(VarName, VarType, VarDefault)					\
	DECLARE_VARIABLE	(##VarName, ##VarType, VarDefault)		\
	STORAGE_REGISTRY	(##VarName, ##VarType)							\
//This macro is used to create a variable not String in shared memory

#define DECLARE_SHARED_MEM_ARRAY_INDEXED(VarName, Size, Index, VarDefault)	\
	ENCAPSULATE_INDEXED_ARRAY(##VarName, char, ##Size, ##Index, ##VarDefault)		\
	DECLARE_VARIABLE_INDEX(##VarName, ST_##VarName, ##Index, st##VarName##default)			\
	STORAGE_SHARED_MEM_INDEX(##VarName, ST_##VarName, ##Index)

#define DECLARE_REGISTRY_ARRAY_INDEXED(VarName, Size, Index, VarDefault)	\
	ENCAPSULATE_INDEXED_ARRAY(##VarName, char, ##Size, ##Index, ##VarDefault)		\
	DECLARE_VARIABLE_INDEX(##VarName, ST_##VarName, ##Index, st##VarName##default)			\
	STORAGE_REGISTRY_INDEX(##VarName, ST_##VarName, ##Index)

#define ENCAPSULATE_INDEXED_ARRAY(VarName, VarType, Size, Index, VarDefault)		\
	struct ST_##VarName												\
	{																\
		VarType	VarName[Size];										\
	} st##VarName##default = { VarDefault };

#define DECLARE_SHARED_MEM_VAR_INDEXED(VarName, VarType, Index, VarDefault)	\
	DECLARE_VARIABLE_INDEX(##VarName, ##VarType, ##Index, ##VarDefault)		\
	STORAGE_SHARED_MEM_INDEX(##VarName, ##VarType, ##Index)

#define DECLARE_REGISTRY_VAR_INDEXED(VarName, VarType, Index, VarDefault)	\
	DECLARE_VARIABLE_INDEX(##VarName, ##VarType, ##Index, ##VarDefault)		\
	STORAGE_REGISTRY_INDEX(##VarName, ##VarType, ##Index)

#define DECLARE_VARIABLE_INDEX(VarName, VarType, Index, VarDefault)	\
																			\
	extern	short	ISerialize##VarName(LPVOID, int Ind, BOOL);						\
																			\
	DL_EXPORT	short	DLIGet##VarName(LPVOID Destination, int Ind)					\
	{																		\
		return ISerialize##VarName(Destination, Ind, TRUE);						\
	}																		\
																			\
	DL_EXPORT	short	DLISet##VarName(LPVOID Value, int Ind)						\
	{																		\
		return ISerialize##VarName(Value, Ind, FALSE);							\
	}																		\
																			\
	DL_EXPORT	int		DLIGetSize##VarName()								\
	{																		\
		return sizeof(VarType);												\
		/*return sizeof(VarType)*Index;*/									\
	}																		\
																			\
	DL_EXPORT	short	DLReset##VarName()									\
	{																		\
		VarType Var=VarDefault;												\
		for(int i= 0 ; i< Index ; ++i)										\
			ISerialize##VarName((LPVOID)&Var, i, FALSE);						\
																			\
		return DL_SUCCESS;													\
	}

#define STORAGE_SHARED_MEM_INDEX(VarName, VarType, Index)					\
																			\
	VarType	VarName[Index]={0};												\
																			\
	short	ISerialize##VarName(LPVOID lpBuffer, int Ind, BOOL bRead)		\
	{																		\
		if((Ind < 0) || (Ind > Index))										\
			return DL_INVALIDINDEX;											\
																			\
		if(bRead)															\
		{																	\
			/* memcpy something */											\
			*(VarType *)lpBuffer=VarName[Ind];								\
		}																	\
		else																\
		{																	\
			/* memcpy something */											\
			VarName[Ind]=*(VarType *)lpBuffer;								\
		}																	\
		return (DL_SUCCESS);												\
	}

#define STORAGE_REGISTRY_INDEX(VarName, VarType, Index)						\
																			\
	VarType	VarName={0};													\
																			\
	short	ISerialize##VarName(LPVOID lpBuffer, int Ind, BOOL bRead)				\
	{																				\
		HKEY hKey;																	\
		LONG Result;																\
		char acSubKey[128];															\
																					\
		if((Ind < 0) || (Ind > Index))												\
			return DL_INVALIDINDEX;													\
																					\
		if(LoadString(hDLinstance, IDS_STRING_SUBKEY, acSubKey, sizeof(acSubKey))==0) \
		{																			\
			return (DL_UNDEFINED);													\
		}																			\
																					\
		Result=RegOpenKey(HKEY_LOCAL_MACHINE, acSubKey, &hKey);						\
		if(ERROR_SUCCESS!=Result)													\
		{																			\
			return (DL_UNDEFINED);													\
		}																			\
																					\
		char stIndex[10] = {0};														\
		itoa(Ind,stIndex ,10);														\
																					\
		LPSTR lpszRegName = new char[strlen(#VarName) + 2 + strlen(stIndex) +1];		\
		strcpy(lpszRegName, #VarName);												\
		strcat(lpszRegName, "[");													\
		strcat(lpszRegName, stIndex);												\
		strcat(lpszRegName, "]");													\
																					\
		if(bRead)																	\
		{																			\
			DWORD dwSize=sizeof(VarType);											\
																					\
			Result=RegQueryValueEx(hKey, lpszRegName, NULL, NULL,					\
									(unsigned char *)lpBuffer, &dwSize);			\
			RegCloseKey(hKey);														\
			if(ERROR_SUCCESS!=Result)												\
			{																		\
				delete []lpszRegName;												\
				return (DL_UNDEFINED);												\
			}																		\
			else																	\
			{																		\
				delete []lpszRegName;												\
				return (DL_SUCCESS);												\
			}																		\
		}																			\
		else																		\
		{																			\
			Result=RegSetValueEx(hKey, lpszRegName, NULL, REG_BINARY,				\
								 (unsigned char *)lpBuffer, sizeof(VarType));		\
			RegCloseKey(hKey);														\
			if(ERROR_SUCCESS!=Result)												\
			{																		\
				delete []lpszRegName;												\
				return (DL_UNDEFINED);												\
			}																		\
			else																	\
			{																		\
				delete []lpszRegName;												\
				return (DL_SUCCESS);												\
			}																		\
		}																			\
																					\
		delete []lpszRegName;														\
	}

#pragma data_seg(".DLSM")
#undef	__DataLink_POST__


#endif  // __DataLink_PRE_EX__