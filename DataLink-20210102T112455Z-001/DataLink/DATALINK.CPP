/******************************************************************************
*                                                                             *
*   DataLink.cpp: Defines the initialization routines for the DLL             *
*                                                                             *
******************************************************************************/

#include "DataLinkGeneral.h"
#include <stdio.h>

// ****************************************************************************
// *
// * Global data
// *
// ****************************************************************************
HINSTANCE	hDLinstance;

char cDLGet		[DL_PREFIX_SIZE];
WORD lDLGet;

char cDLIGet	[DL_PREFIX_SIZE];
WORD lDLIGet;

char cDLFGet	[DL_PREFIX_SIZE];
WORD lDLFGet;

char cDLGetSize	[DL_PREFIX_SIZE];
WORD lDLGetSize;

char cDLIGetSize	[DL_PREFIX_SIZE];
WORD lDLIGetSize;

char cDLSet		[DL_PREFIX_SIZE];
WORD lDLSet;

char cDLISet	[DL_PREFIX_SIZE];
WORD lDLISet;

char cDLFSet	[DL_PREFIX_SIZE];
WORD lDLFSet;

char cDLReset	[DL_PREFIX_SIZE];
WORD lDLReset;

char cDLWait	[DL_PREFIX_SIZE];
WORD lDLWait;


// ****************************************************************************
// *
// * Initialization/termination routine
// *
// ****************************************************************************
#ifdef __cplusplus
extern "C"
{
#endif

short GetNameTypeAndIndex(LPSTR Test, LPSTR* lppVarName, int* lpIndex);

BOOL WINAPI
DllMain(HANDLE hInst, ULONG ul_reason_for_call, LPVOID lpReserved)
{
	switch (ul_reason_for_call)
	{
		case    DLL_PROCESS_ATTACH:     // 1
		{
			int i;

			// * Save hInstance for further uses
			hDLinstance = (HINSTANCE)hInst;

			// * Load functions prefix
			// * DLGet
			i = LoadString (hDLinstance,IDS_STRING_DLGET,cDLGet,DL_PREFIX_SIZE);
			cDLGet[i] = 0;
			lDLGet = strlen (cDLGet);

			// * DLIGet
			i = LoadString (hDLinstance,IDS_STRING_DLIGET,cDLIGet,DL_PREFIX_SIZE);
			cDLIGet[i] = 0;
			lDLIGet = strlen (cDLIGet);

			// * DLSet
			i = LoadString (hDLinstance,IDS_STRING_DLSET,cDLSet,DL_PREFIX_SIZE);
			cDLSet[i] = 0;
			lDLSet = strlen (cDLSet);

			// * DLISet
			i = LoadString (hDLinstance,IDS_STRING_DLISET,cDLISet,DL_PREFIX_SIZE);
			cDLISet[i] = 0;
			lDLISet = strlen (cDLISet);

			// * DLGetSize
			i = LoadString (hDLinstance,IDS_STRING_DLGETSIZE,cDLGetSize,DL_PREFIX_SIZE);
			cDLGetSize[i] = 0;
			lDLGetSize = strlen (cDLGetSize);

			// * DLIGetSize
			i = LoadString (hDLinstance,IDS_STRING_DLIGETSIZE,cDLIGetSize,DL_PREFIX_SIZE);
			cDLIGetSize[i] = 0;
			lDLIGetSize = strlen (cDLIGetSize);

			// * DLFGet
			i = LoadString (hDLinstance,IDS_STRING_DLFGET,cDLFGet,DL_PREFIX_SIZE);
			cDLFGet[i] = 0;
			lDLFGet = strlen (cDLFGet);

			// * DLFSet
			i = LoadString (hDLinstance,IDS_STRING_DLFSET,cDLFSet,DL_PREFIX_SIZE);
			cDLFSet[i] = 0;
			lDLFSet = strlen (cDLFSet);

			// * DLReset
			i = LoadString (hDLinstance,IDS_STRING_DLRESET,cDLReset,DL_PREFIX_SIZE);
			cDLReset[i] = 0;
			lDLReset = strlen (cDLReset);

			// * DLWait
			i = LoadString (hDLinstance,IDS_STRING_DLWAIT,cDLWait,DL_PREFIX_SIZE);
			cDLWait[i] = 0;
			lDLWait = strlen (cDLWait);
		}
		break;

		/*
		case    DLL_PROCESS_DETACH:     // 0
		case    DLL_THREAD_ATTACH:      // 2
		case    DLL_THREAD_DETACH:      // 3
		*/
		default:
		break;
	}
	return (TRUE);
}


void __cdecl
Trace(const char *fmt, ...)
{
    va_list     vaArgs;
    char        outBuffer [1024];

    va_start (vaArgs, fmt);
    vsprintf (outBuffer, fmt, vaArgs);
    OutputDebugString (outBuffer);

    return;
}

////////////////////////////////////////////////////////////////////////////
//
// Get / Set     generic data access
//
////////////////////////////////////////////////////////////////////////////
DL_DEFINE	short
DLGet(LPSTR VarName, LPVOID Destination, LPSTR Format)
{
	char        ProcName[128];
	LPFNGet     lpfnGet = NULL;
	LPFNIGet    lpfnIGet = NULL;
	LPFNFGet    lpfnFGet = NULL;

	ProcName[0] = 0;
	if(Format)
	{
		ProcName[0] ='_';
		strcpy (&ProcName[1],			cDLFGet);
		strcpy (&ProcName[lDLFGet+1],	VarName);
		strcat (ProcName,				"@8");
		lpfnFGet = (LPFNFGet) GetProcAddress (hDLinstance, ProcName);
		if(lpfnFGet)
		{
			return (lpfnFGet(Destination, Format));
		}
	}
	else
	{
		short Action;
		LPSTR VarIndexName = NULL;
		LPSTR RealVarName = NULL;
		int Index = 0;

		Action = GetNameTypeAndIndex(VarName, &VarIndexName, &Index);
		
		ProcName[0] ='_';

		switch(Action)
		{
			case 2:
				return DL_UNDEFINED;
			case 0:
				RealVarName = VarName;
				strcpy (&ProcName[1], cDLGet);
				strcpy (&ProcName[lDLGet+1], RealVarName);
				strcat (ProcName, "@4");
				lpfnGet = (LPFNGet) GetProcAddress (hDLinstance, ProcName);
				break;
			case 1:
				RealVarName = VarIndexName;
				strcpy (&ProcName[1], cDLIGet);
				strcpy (&ProcName[lDLIGet+1], RealVarName);
				strcat (ProcName, "@8");
				lpfnIGet = (LPFNIGet) GetProcAddress (hDLinstance, ProcName);
				break;
		}

		if(VarIndexName)
			delete [] VarIndexName;

		if(lpfnGet)
		{
			return (lpfnGet(Destination));
		}

		if(lpfnIGet)
		{
			return (lpfnIGet(Destination, Index));
		}

	}
	return (DL_UNDEFINED);
}


DL_DEFINE	short
DLSet(LPSTR VarName, LPVOID Value, LPSTR Format)
{
	char        ProcName[128];
	LPFNSet     lpfnSet = NULL;
	LPFNISet    lpfnISet = NULL;
	LPFNFGet    lpfnFSet = NULL;

	ProcName[0] = 0;
	if(Format)
	{
		ProcName[0] ='_';
		strcpy (&ProcName[1],			cDLFSet);
		strcpy (&ProcName[lDLFSet+1],	VarName);
		strcat (ProcName,				"@8");
		lpfnFSet = (LPFNFSet) GetProcAddress (hDLinstance, ProcName);
		if(lpfnFSet)
		{
			return (lpfnFSet(Value, Format));
		}
	}
	else
	{
		short Action;
		LPSTR VarIndexName = NULL;
		LPSTR RealVarName = NULL;
		int Index = 0;

		Action = GetNameTypeAndIndex(VarName, &VarIndexName, &Index);

		ProcName[0] ='_';

		switch(Action)
		{
			case 2:
				return DL_UNDEFINED;
			case 0:
				RealVarName = VarName;
				strcpy (&ProcName[1], cDLSet);
				strcpy (&ProcName[lDLSet+1], RealVarName);
				strcat (ProcName, "@4");
				lpfnSet = (LPFNSet) GetProcAddress (hDLinstance, ProcName);
				break;
			case 1:
				RealVarName = VarIndexName;
				strcpy (&ProcName[1], cDLISet);
				strcpy (&ProcName[lDLISet+1], RealVarName);
				strcat (ProcName, "@8");
				lpfnISet = (LPFNISet) GetProcAddress (hDLinstance, ProcName);
				break;
		}

		if(VarIndexName)
			delete [] VarIndexName;

		if(lpfnSet)
		{
			return (lpfnSet(Value));
		}

		if(lpfnISet)
		{
			return (lpfnISet(Value, Index));
		}
	}
	return (DL_UNDEFINED);
}

////////////////////////////////////////////////////////////////////////////
//
// GetSize
//
////////////////////////////////////////////////////////////////////////////
DL_DEFINE	int
DLGetSize(LPSTR VarName)
{
	char		ProcName[128];
	LPFNGetSize		lpfnGetSize = NULL;
	LPFNIGetSize	lpfnIGetSize = NULL;

	short Action;
	LPSTR VarIndexName = NULL;
	LPSTR RealVarName = NULL;
	int Index = 0;

	Action = GetNameTypeAndIndex(VarName, &VarIndexName, &Index);

	ProcName[0] ='_';

	switch(Action)
	{
		case 2:
			return DL_UNDEFINED;
		case 0:
			RealVarName = VarName;
			strcpy (&ProcName[1], cDLGetSize);
			strcpy (&ProcName[lDLGetSize+1], RealVarName);
			strcat (ProcName, "@0");
			lpfnGetSize = (LPFNGetSize) GetProcAddress (hDLinstance, ProcName);
			break;
		case 1:
			RealVarName = VarIndexName;
			strcpy (&ProcName[1], cDLIGetSize);
			strcpy (&ProcName[lDLIGetSize+1], RealVarName);
			strcat (ProcName, "@0");
			lpfnIGetSize = (LPFNIGetSize) GetProcAddress (hDLinstance, ProcName);
			break;
	}

	if(VarIndexName)
		delete [] VarIndexName;

	if(lpfnGetSize)
	{
		return (lpfnGetSize());
	}
	
	if(lpfnIGetSize)
	{
		return (lpfnIGetSize());
	}
	return (DL_UNDEFINED);
}

////////////////////////////////////////////////////////////////////////////
//
// Wait
//
////////////////////////////////////////////////////////////////////////////
DL_DEFINE	DWORD
DLWait(LPSTR VarName, int nTimeout)
{
	char        ProcName[128];
	LPFNWait	lpfnWait;


	ProcName[0] ='_';
	strcpy (&ProcName[1],			cDLWait);
	strcpy (&ProcName[lDLWait+1],	VarName);
	strcat (ProcName,				"@4");

	lpfnWait = (LPFNWait) GetProcAddress (hDLinstance, ProcName);
	if (lpfnWait)
	{
		return (lpfnWait(nTimeout));
	}
	else
	{
		return (DL_UNDEFINED);
	}
}

////////////////////////////////////////////////////////////////////////////
//
// Reset
//
////////////////////////////////////////////////////////////////////////////
DL_DEFINE	short
DLReset(LPSTR VarName)
{
	char        ProcName[128];
	LPFNReset   lpfnReset;


	ProcName[0] ='_';
	strcpy (&ProcName[1],			cDLReset);
	strcpy (&ProcName[lDLReset+1],	VarName);
	strcat (ProcName,				"@0");

	lpfnReset = (LPFNReset) GetProcAddress (hDLinstance, ProcName);
	if (lpfnReset)
	{
		return (lpfnReset());
	}
	else
	{
		return (DL_UNDEFINED);
	}
}


DL_DEFINE	short
DLGReset(LPSTR GroupName)
{
	char acGroupKey[128];
	char *pcVariables=NULL;
	char *pcTmp=NULL;
	DWORD dwSize=0;
	int nCounter=0;
	HKEY hKey=NULL;
	short nReply=DL_SUCCESS;

	// Load the string resource that contains the name of the reg key to open
	if(LoadString(hDLinstance, IDS_STRING_GROUPKEY, acGroupKey, sizeof(acGroupKey))==0)
	{
		return (DL_UNDEFINED);
	}
	
	// Open the key
	if(RegOpenKey(HKEY_LOCAL_MACHINE, acGroupKey, &hKey)!=ERROR_SUCCESS)
	{
		return (DL_UNDEFINED);
	}

	// Retrieve the size of the value of the group we want to work with
	if(RegQueryValueEx(hKey, GroupName, NULL, NULL, NULL, &dwSize)!=ERROR_SUCCESS)
	{
		RegCloseKey(hKey);
		return (DL_UNDEFINED);
	}

	// Allocate enough storage to load the group
	pcVariables=new char[dwSize];

	// Load the group
	if(RegQueryValueEx(hKey, GroupName, NULL, NULL, (unsigned char *)pcVariables, 
						&dwSize)!=ERROR_SUCCESS)
	{
		RegCloseKey(hKey);
		delete[] pcVariables;
		return (DL_UNDEFINED);
	}

	// Walk through all the variables in the group
	pcTmp=pcVariables;
	// Use a counter to provide some kind of feedback on which variable caused a problem
	nCounter=1;
	// Now ch
	while(*pcTmp)
	{
		if(DLReset(pcTmp)!=DL_SUCCESS)
		{
			// If DLReset fails, then we return an error indicating that some
			// error occurred on a variable
			// A more specific error reporting mechanism has to be implemented!!
			nReply=DL_GRESET_VARIABLE_ERROR;
		}
		pcTmp+=strlen(pcTmp)+1;
		nCounter++;
	}

	delete[] pcVariables;

	return (nReply);
}

short GetNameTypeAndIndex(LPSTR Test, LPSTR* lppVarName, int* lpIndex)
{
	if(!Test)
		return 2;

	LPSTR NameStart= NULL, IndexStart = NULL, LastStart = NULL;
	LPSTR TmpTest = new char[strlen(Test)+1];
	LPSTR TmpTestCheck = new char[strlen(Test)+3];
	LPSTR StrIndex = new char[strlen(Test)+1];
	int Index= -1;
	short Action = 0; //0 == Not Indexed 1 == Indexed 2 == Error
	
	memset(TmpTest, 0x00, strlen(Test)+1);
	memset(TmpTestCheck, 0x00, strlen(Test)+1);
	memset(StrIndex, 0x00, strlen(Test)+1);
	memcpy(TmpTest, Test, strlen(Test));

	NameStart = TmpTest;

	IndexStart = strstr(TmpTest, "[");
	
	if(IndexStart != NULL) //Potrebbe essere un Index
	{
		*IndexStart = 0;
		IndexStart++;

		LastStart = strstr(IndexStart, "]");
		if(LastStart)
		{
			*LastStart = 0;
			LastStart++;

			Index = atoi(IndexStart);
			itoa(Index, StrIndex, 10);
		}
	}
	
	strcpy(TmpTestCheck, NameStart);
	strcat(TmpTestCheck, "[");
	strcat(TmpTestCheck, StrIndex);
	strcat(TmpTestCheck, "]");
	//strcat(TmpTestCheck, LastStart);

	if(!IndexStart)
		Action = 0;
	else
		Action = (strcmp(TmpTestCheck, Test) == 0 ? 1 : 2);

	switch(Action)
	{
	case 1:
		*lpIndex = Index;
		*lppVarName = new char [strlen(NameStart)+1];
		strcpy(*lppVarName, NameStart);
		break;
	case 0:
	case 3:
		break;
	}

	if(TmpTest)
		delete [] TmpTest;
	if(TmpTestCheck)
		delete [] TmpTestCheck;
	if(StrIndex)
		delete [] StrIndex;

	return Action;
}


#ifdef __cplusplus
}
#endif
