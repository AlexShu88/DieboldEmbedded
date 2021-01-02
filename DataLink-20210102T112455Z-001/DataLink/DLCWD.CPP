/******************************************************************************
*                                                                             *
*  DlCwd.cpp: Defines Withdraw variables                                     *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

/////////////////////////////////////////////////////////////////////////////
/** Withdraw Amount input by customer **/
	DECLARE_SHARED_MEM_ARRAY(CwdAmount,		10,	"")

// Available minimum denomination
	DECLARE_SHARED_MEM_ARRAY(CwdAvailDenom,	4,	"")

// Withdraw limited amount. the mininum of FitMaxWthAmount and the 
// balance of current account 
	DECLARE_SHARED_MEM_ARRAY(CwdWthLimited,	8,	"")

//The trigger is to indicate whether CWD Crimre Possible
DECLARE_REGISTRY_ARRAY( CWDCrimePossible,	2,	"" ) 

/////////////////////////////////////////////////////////////////////////////

#include "DataLinkPost.h"