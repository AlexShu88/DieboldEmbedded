/******************************************************************************
*                                                                             *
* DlCashIn.cpp: Defines CashIn transaction variables                          *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

// The total number of successful CashIn.
DECLARE_REGISTRY_INTEGER( TotCashInNum,		0 )

// The total amount of CashIn.
DECLARE_REGISTRY_DOUBLE( TotCashInAmount,		0.0 )

// The total number of Note CNY10  
DECLARE_REGISTRY_INTEGER( TotCashInNum10,			0)

// The total number of Note CNY20  
DECLARE_REGISTRY_INTEGER( TotCashInNum20,			0)

// The total number of Note CNY50  
DECLARE_REGISTRY_INTEGER( TotCashInNum50,			0)

// The total number of Note CNY100  
DECLARE_REGISTRY_INTEGER( TotCashInNum100,			0)


// The number of Note CNY50 during cashIn transaction 
DECLARE_SHARED_MEM_INT( CashInNum10,			0)

// The number of Note CNY100 during cashIn transaction 
DECLARE_SHARED_MEM_INT( CashInNum20,			0)

// The number of Note CNY50 during cashIn transaction 
DECLARE_SHARED_MEM_INT( CashInNum50,			0)

// The number of Note CNY100 during cashIn transaction 
DECLARE_SHARED_MEM_INT( CashInNum100,			0)

// The number of NotesRefused during cashIn transaction 
DECLARE_SHARED_MEM_INT( CashInNotesRefused,	0)

// The number of Note CNY50 during cashIn transaction 
DECLARE_SHARED_MEM_ARRAY( CashInNum10Prr,	5,	"")

// The number of Note CNY50 during cashIn transaction 
DECLARE_SHARED_MEM_ARRAY( CashInNum20Prr,	5,	"")

// The number of Note CNY50 during cashIn transaction 
DECLARE_SHARED_MEM_ARRAY( CashInNum50Prr,	5,	"")

// The number of Note CNY100 during cashIn transaction 
DECLARE_SHARED_MEM_ARRAY( CashInNum100Prr,	5,	"")

// The Maxnumber of cashIn transaction 
DECLARE_SHARED_MEM_ARRAY(CashInMaxNum,		4,	"" )

// The Accepted Denomination of cashIn transaction 
DECLARE_SHARED_MEM_ARRAY(CashInAcceptedDenomination,		20,	"" )

// Html Screen CashInPrompt10
DECLARE_SHARED_MEM_ARRAY( HtmlCashInPrompt10,	40,	"")

// Html Screen CashInPrompt20
DECLARE_SHARED_MEM_ARRAY( HtmlCashInPrompt20,	40,	"")

// Html Screen CashInPrompt50
DECLARE_SHARED_MEM_ARRAY( HtmlCashInPrompt50,	40,	"")

// Html Screen CashInPrompt100
DECLARE_SHARED_MEM_ARRAY( HtmlCashInPrompt100,	40,	"")
	


//DECLARE_SHARED_MEM_ARRAY( CashInIsTakeNoteTimeout,	2,	"N" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY11,		4,	"11" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY12,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY13,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY14,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY15,		6,	"" )
	
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY21,		4,	"12" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY22,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY23,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY24,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY25,		6,	"" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY31,		4,	"13" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY32,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY33,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY34,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY35,		6,	"" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY41,		4,	"14" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY42,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY43,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY44,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY45,		6,	"" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY51,		4,	"15" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY52,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY53,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY54,		6,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWorkCNY55,		6,	"" )

// Individual account identifier in Track2 or Track3
DECLARE_SHARED_MEM_ARRAY( CIMRecoveryAccNo,				25,	"" )

#include "DataLinkPost.h"