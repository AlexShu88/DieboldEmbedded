/******************************************************************************
*                                                                             *
*   DataLinkGeneral.h:   main header file for the DataLink DLL                *
*                                                                             *
******************************************************************************/



#ifndef __DataLinkGeneral_INC__

    #define __DataLinkGeneral_INC__


    #include <windows.h>


	////////////////////////////////////////////////////////////////////////////
	//
	// Function prototypes
	//
	////////////////////////////////////////////////////////////////////////////
	#define DL_DEFINE		__declspec(dllexport)
	#include "DataLink.h"


	////////////////////////////////////////////////////////////////////////////
	//
	// Constants
	//
	////////////////////////////////////////////////////////////////////////////
	#include "DLConstants.h"


    ////////////////////////////////////////////////////////////////////////////
    //
    // Function types
    //
    ////////////////////////////////////////////////////////////////////////////
    typedef short   (*LPFNGet)		(LPVOID);
	typedef short   (*LPFNIGet)		(LPVOID, int);
    typedef short   (*LPFNFGet)		(LPVOID, LPSTR);
    typedef short   (*LPFNSet)		(LPVOID);
	typedef short   (*LPFNISet)		(LPVOID, int);
    typedef short   (*LPFNFSet)		(LPVOID, LPSTR);
    typedef short   (*LPFNReset)	();
	typedef int		(*LPFNGetSize)	();
	typedef int		(*LPFNIGetSize)	();
	typedef DWORD	(*LPFNWait)		(int);



#endif  // __DataLinkGeneral_INC__
