/******************************************************************************
*                                                                             *
*  DLPrrFields.cpp: Defines Fields for printer on Reciept printer             *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

/////////////////////////////////////////////////////////////////////////////
// Define the formatted Trnasaction amount field 
DECLARE_SHARED_MEM_ARRAY(PrrTransAmount,	20,	"")

// Define the transfer 2nd account printed on receipt for CMB Shenzhen  
DECLARE_SHARED_MEM_ARRAY(PrrTfr2ndAccNo,	24,	"")

// Define the formatted account balance field 
DECLARE_SHARED_MEM_ARRAY(PrrAccBalance,		18,	"")

// Define the star mark of "CashIn" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrCashInMark,	4,	"")

// Define the star mark of "Withdraw" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrWthMark,	4,	"")

// Define the star mark of "Transfer" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrTransferMark,	4,	"")

// Define the star mark of "Others" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrOthersMark,	4,	"")

// Define the star mark of "Retain card" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrCardRetainMark,	4,	"")

// Define the star mark of "Please contact with bank" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrContactBankMark,	4,	"")

// Define the star mark of "Accept" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrAcceptMark,	4,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRejectMark,	4,	"")

DECLARE_SHARED_MEM_ARRAY(PrrAcceptCode,	7,	"")

DECLARE_SHARED_MEM_ARRAY(PrrRejectedCode,	4,	"")

// Define trans code for writing capture card log for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrCpdTransCode,	3,	"")

// Define the star mark of "Transcation Type" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrTransType,	8,	"")

// Define the star mark of "Fee Charge" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrFeeCharge,	40,	"")

// Define the star mark of "Host Enq-No" field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY(PrrHostEnqNo,	44,	"")

/////////////////////////////////////////////////////////////////////////////

// Define buffer for printing Chinese  (BOC shanghai)
DECLARE_SHARED_MEM_ARRAY(PrrRow1,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow2,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow3,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow4,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow5,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow6,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow7,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow8,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow9,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow10,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow11,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow12,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow13,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow14,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow15,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow16,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow17,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow18,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow19,	40,	"")
DECLARE_SHARED_MEM_ARRAY(PrrRow20,	40,	"")

#include "DataLinkPost.h"