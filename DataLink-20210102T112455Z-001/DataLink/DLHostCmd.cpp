/******************************************************************************
*                                                                             *
*  DlHostCmd.cpp: Defines Host command variables                              *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

/////////////////////////////////////////////////////////////////////////////
// CDM Cassette1 Initial num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassIniNum1,		6,	"00000")

// CDM Cassette2 Initial num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassIniNum2,		6,	"00000")

// CDM Cassette3 Initial num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassIniNum3,		6,	"00000")

// CDM Cassette4 Initial num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassIniNum4,		6,	"00000")

// CDM Cassette1 Left num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassLeftNum1,	6,	"00000")

// CDM Cassette2 Left num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassLeftNum2,	6,	"00000")

// CDM Cassette3 Left num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassLeftNum3,	6,	"00000")

// CDM Cassette4 Left num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassLeftNum4,	6,	"00000")

// CDM Cassette1 Reject num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassRejectNum1,	6,	"00000")

// CDM Cassette2 Reject num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassRejectNum2,	6,	"00000")

// CDM Cassette3 Reject num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassRejectNum3,	6,	"00000")

// CDM Cassette4 Reject num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCDMCassRejectNum4,	6,	"00000")

// CIM CNY5 CashIn num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCIMCNY5,			6,	"00000")

// CIM CNY10 CashIn num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCIMCNY10,			6,	"00000")

// CIM CNY20 CashIn num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCIMCNY20,			6,	"00000")

// CIM CNY50 CashIn num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCIMCNY50,			6,	"00000")

// CIM CNY100 CashIn num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCIMCNY100,			6,	"00000")

//  Total Withdraw num
	DECLARE_SHARED_MEM_ARRAY(HostCmdWithdrawNum,	    6,	"00000")
//  Total Withdraw amount
	DECLARE_SHARED_MEM_ARRAY(HostCmdWithdrawAmount,		15,	"000000000000")

//  Total cashin num
	DECLARE_SHARED_MEM_ARRAY(HostCmdCashInNum,	        6,	"00000")

//  Total CashIn amount
	DECLARE_SHARED_MEM_ARRAY(HostCmdCashInAmount,		15,	"000000000000")

//  Total transfer num
	DECLARE_SHARED_MEM_ARRAY(HostCmdTfrOutNum,	         6,	"00000")

//  Total Transfer out amount
	DECLARE_SHARED_MEM_ARRAY(HostCmdTfrOutAmount,		15,	"000000000000")

//  Total Withdraw reversal num
	DECLARE_SHARED_MEM_ARRAY(HostCmdWthRevNum,	    6,	"00000")

//  Total withdraw reversal amount
//  Add for Icbc3030 project
	DECLARE_SHARED_MEM_ARRAY(HostCmdWthRevAmount,		15,	"000000000000")

//  Total list num
	DECLARE_SHARED_MEM_ARRAY(HostCmdListNum,	        6,	"00000")

//  Total Withdraw amount for extra bank
//  Add for Icbc3030 project
	DECLARE_SHARED_MEM_ARRAY(HostCmdExtraWthAmount,		15,	"000000000000")

//  Total withdraw reversal amount for extra bank
//  Add for Icbc3030 project
	DECLARE_SHARED_MEM_ARRAY(HostCmdExtraWthRevAmount,		15,	"000000000000")
/////////////////////////////////////////////////////////////////////////////

#include "DataLinkPost.h"