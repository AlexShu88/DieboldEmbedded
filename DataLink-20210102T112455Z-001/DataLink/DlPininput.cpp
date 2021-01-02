/******************************************************************************
*                                                                             *
*  DlPininput.cpp: Defines Pininput variables                                     *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

// The encrypted pin, the Pin Block 
	DECLARE_SHARED_MEM_ARRAY(PinInputBlock,		18,	"")

// The customer input pin field.
	DECLARE_SHARED_MEM_ARRAY(PinInputPin,		8,	"")

#include "DataLinkPost.h"