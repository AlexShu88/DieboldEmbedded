/******************************************************************************
*                                                                             *
*  DlTransfer.cpp: Defines TransferIn and TransferOut variables                                     *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

/////////////////////////////////////////////////////////////////////////////
// Keeping the account number of Another card which the customer
// want to transfer to
DECLARE_SHARED_MEM_ARRAY(Tfr2ndAccNo,	24,	"")

//Indicate the transfer type
DECLARE_SHARED_MEM_ARRAY(TfrType,		6,	"")

/////////////////////////////////////////////////////////////////////////////

#include "DataLinkPost.h"