/******************************************************************************
*                                                                             *
*  HostReturn.cpp: Defines variables return from Host                                     *
*                                                                             *
******************************************************************************/
#include "DataLinkPre.h"

// Keep the reject code field in Host rejection message
DECLARE_SHARED_MEM_ARRAY(HostRejectCode,	5, "")

// The reason of rejection in Chinese character, 
// for displaying on customer screen
DECLARE_SHARED_MEM_ARRAY(HostRejectChinese,		60,	"")

// The reason of rejection in English character, 
// for printing on journal
DECLARE_SHARED_MEM_ARRAY(HostRejectEnglish,		120,	"")

// Indicate whether this Host rejection code should capture the 
// card or not? "Y" - Capture; "N" - Not Capture
DECLARE_SHARED_MEM_ARRAY(HostRejectCard,		2,	"")

// Indicate the Stocking account balance
DECLARE_SHARED_MEM_ARRAY(HostSerialNo,		21, "")

// Keep the Track2 message from Host
DECLARE_SHARED_MEM_ARRAY(HostTrack2,	40, "")

// Keep the Track3 message from Host
DECLARE_SHARED_MEM_ARRAY(HostTrack3,	110, "")

// Keep host return Account Number for check
DECLARE_SHARED_MEM_ARRAY(HostAccNo,			24, "")

// Keep host return atmcode for check
DECLARE_SHARED_MEM_ARRAY(HostAtmCode,		5, "")

// Keep host return trans. code for verify
DECLARE_SHARED_MEM_ARRAY(HostTransCode,		4, "")

DECLARE_SHARED_MEM_ARRAY(HostLineNum,		6, "")

DECLARE_SHARED_MEM_ARRAY(HostFlagCode,		5, "")

DECLARE_SHARED_MEM_ARRAY(HostTransAmount,	9, "")

//For ICBC_HQ Begin Add by lijun
	DECLARE_SHARED_MEM_ARRAY(HostAdtRecCnt,		4, "")
	DECLARE_SHARED_MEM_ARRAY(HostAnqPkgCnt,		4, "")
	DECLARE_SHARED_MEM_ARRAY(HostCurBal,		12, "")
	DECLARE_SHARED_MEM_ARRAY(HostCurrentDate,	13, "")
    DECLARE_SHARED_MEM_ARRAY(HostCurrentTime,	7, "")
	DECLARE_SHARED_MEM_ARRAY(HostFundAvail,		9, "")
	DECLARE_SHARED_MEM_ARRAY(HostLimitCode,		2, "")
	DECLARE_SHARED_MEM_ARRAY(HostLimitFlg,		2, "")
	DECLARE_SHARED_MEM_ARRAY(HostShAvailBal,	12, "")
	DECLARE_SHARED_MEM_ARRAY(HostShTfrAvail,	9, "")
	DECLARE_SHARED_MEM_ARRAY(HostSignAvailBal,	2, "")
	DECLARE_SHARED_MEM_ARRAY(HostSignCurBal,	2, "")
	DECLARE_SHARED_MEM_ARRAY(HostTransAvailTimes,2, "")

	DECLARE_SHARED_MEM_ARRAY(VdrRcCode,			5, "")
   // Keep the reject code field in ATMP rejection message
    DECLARE_SHARED_MEM_ARRAY(ATMPRejectCode, 5, "")
//For ICBC_HQ End Add

#include "DataLinkPost.h"
