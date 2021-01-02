/******************************************************************************
*                                                                             *
*  DLOpteva.cpp: Defines new variables for Opteva                             *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

// Keep the current value of monitor type. 
// The trigger to inform program of the changed of monitor type
DECLARE_WAITABLE_SHARED_MEM_INT( OptevaMonType, 0 )

DECLARE_SHARED_MEM_ARRAY( OptevaCasStatus,		1024,	"" )

DECLARE_SHARED_MEM_ARRAY( GBLFld1,55,	"" )

DECLARE_SHARED_MEM_ARRAY( GBLFld2,55,	"" )

DECLARE_SHARED_MEM_ARRAY( GBLFld3,55,	"" )

DECLARE_SHARED_MEM_ARRAY( GBLFld4,55,	"" )

DECLARE_SHARED_MEM_ARRAY( GBLFld5,55,	"" )    

#include "DataLinkPost.h"