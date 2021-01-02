/******************************************************************************
*                                                                             *
* DlPinChange.cpp: Defines Pin change transaction variables                            *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"


// New Pin First input field
DECLARE_SHARED_MEM_ARRAY( PinChangeInput1,		8,	"")

// New Pin Second input field
DECLARE_SHARED_MEM_ARRAY( PinChangeInput2,		8,	"")

// Error Prompt in PinChangeDiff.htm file
DECLARE_SHARED_MEM_ARRAY( PinChangeDiff,		40,	"")

// Error Prompt in PinChangeDiff.htm file
DECLARE_SHARED_MEM_ARRAY( PinChangeRetry,		2,	"")

//Error Prompt in PinChangeError.htm file
DECLARE_SHARED_MEM_ARRAY( PinChangeError,		40,	"")

// The encrypted new pin, the New PIN block
DECLARE_SHARED_MEM_ARRAY( PinChangeBlock,		18,	"")

#include "DataLinkPost.h"


