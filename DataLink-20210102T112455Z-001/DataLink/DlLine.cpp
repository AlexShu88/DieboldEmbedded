#include "DataLinkPre.h"
	DECLARE_REGISTRY_INTEGER(MessNumber, 0);

// The host message sequeue number.
// Increasing for each communication message
	DECLARE_REGISTRY_INTEGER(LineSeqNumber, 0);

//use as comunitcation sequence number
	DECLARE_SHARED_MEM_ARRAY(LineLuNumber, 5 ,"")
	//Lu number,copy from the last 3  of ATM code in s3estart

#include "DataLinkPost.h"
