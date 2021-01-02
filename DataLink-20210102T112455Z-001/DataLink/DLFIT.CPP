/******************************************************************************
*                                                                             *
*   DLfit.cpp: Defines the variables access from fit.ini                             *
*                                                                             *
******************************************************************************/

#include <windows.h>

#include "DataLinkPre.h"

// Track3 Message
DECLARE_SHARED_MEM_ARRAY( FitTrack3Message,		120,	"" )

// Track2 Message
DECLARE_SHARED_MEM_ARRAY( FitTrack2Message,		120,	"" )

//Maximum withdraw amount limit for each time
DECLARE_SHARED_MEM_ARRAY( FitMaxWthAmount,		7,	"" )

//For ICBC_HQ Begin Add by lijun
//Pin input length
DECLARE_SHARED_MEM_ARRAY( FitCardPinLength,		2,	"" )

//Pin max attempt.  Used fro recover the pin retry times on Track3
DECLARE_SHARED_MEM_ARRAY( FitPinMaxAttempt,		2,	"" )
//For ICBC_HQ End Add

// Individual account identifier in Track2 or Track3
DECLARE_SHARED_MEM_ARRAY( FitAccNo,				25,	"" )

// Individual account identifier in Track2 or Track3, for Receipt printing
DECLARE_SHARED_MEM_ARRAY( FitPrrAccNo,			25,	"" )

// handle dpkey for 0-abc saving card; 1-abc credit card; 2-other bank's saving card;3-other bank's credit card; 9-visa card
DECLARE_SHARED_MEM_ARRAY( FitCardType, 		    5,	"")

#include "DataLinkPost.h"
