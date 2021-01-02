/******************************************************************************
*                                                                             *
*  DlTotals.cpp: Defines Totals variables                                     *
*                                                                             *
******************************************************************************/
#include "DataLinkPre.h"

/********************************************************************
*																	*
*		Define the system status total fields						*
*																	*
********************************************************************/
//The date and time of the last period closed
DECLARE_REGISTRY_ARRAY( TotPeriodCloseTime,		21,		"" )

// The date and time of the last period opened
// Format YYYY/MM/DD HH:MM:SS
DECLARE_REGISTRY_ARRAY( TotPeriodOpenTime,		21,		"" )

//The date and time of the last loading bank notes.
// Format YYYY/MM/DD HH:MM:SS
DECLARE_REGISTRY_ARRAY( TotLoadNoteTime,		21,		"" )
/*********************************************************************/

/********************************************************************
*																	*
*		Define the transactions statistic total fields				*
*																	*
********************************************************************/
// The total number of successful withdrawals.
DECLARE_REGISTRY_INTEGER( TotWithdrawNum,		0 )

// The total amount of withdrawal.
DECLARE_REGISTRY_DOUBLE( TotWithdrawAmount,		0.0 )

// The total number of successful transfer out transactions.
DECLARE_REGISTRY_INTEGER( TotTfrOutNum,			0 )

// The total amount of transfer out.
DECLARE_REGISTRY_DOUBLE( TotTfrOutAmount,		0.0 )

// The total number of Inquiry transactions.
DECLARE_REGISTRY_INTEGER( TotInquiryNum,		0 )

// The total number of Pinchange transactions.
DECLARE_REGISTRY_INTEGER( TotPinChangeNum,		0 )

// The total number of captured cards.
DECLARE_REGISTRY_INTEGER( TotCapCardNum,		0 )

// The sequence number of journal printer forms
DECLARE_REGISTRY_INTEGER( TotJournalNum,		0 )

// Add for Icbc3030 project
DECLARE_REGISTRY_INTEGER( TotWthReversalNum,		0 )

// The total amount of withdrawal reversal.
DECLARE_REGISTRY_DOUBLE( TotWthReversalAmount,		0.0 )


#include "DataLinkPost.h"