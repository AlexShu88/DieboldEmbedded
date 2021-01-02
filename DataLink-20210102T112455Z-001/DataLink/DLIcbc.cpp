/******************************************************************************
*                                                                             *
* DLBOC.cpp: Defines new variables for Bank of China Opteva project Sep. 2005 *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

// The balance of available withdraw amount today for BOC
DECLARE_SHARED_MEM_ARRAY(HtmlAvailWTHBalance,		20,	"")

// The balance of available transfer amount today for BOC
DECLARE_SHARED_MEM_ARRAY(HtmlAvailTFRBalance,		20,	"")

// A Incorrect PIN flag for showing in PinInputPin.htm for BOC
DECLARE_SHARED_MEM_ARRAY(GBLPwdIsWrong,		2,	"N")

// Host response Date&Time used for CWC 	
DECLARE_SHARED_MEM_ARRAY(ICBCHostSettleDate,		16,		"")

// Host response Date&Time in Field48 of "8201" format.	
DECLARE_SHARED_MEM_ARRAY(IcbcHostDateTime,		16,		"")

// The Date&Time version of FIT Table downloaded in field48 of "8201" format
DECLARE_REGISTRY_ARRAY(IcbcFitTableVersion,		16,		"")

// The Date&Time version of others Black list file downloaded in field48 of "8201" format
DECLARE_REGISTRY_ARRAY(IcbcBlackOthVersion,		16,		"")


// The trigger is to indicate whether the A&B keys have changed
DECLARE_SHARED_MEM_ARRAY( IcbcMaxTfrAmount,		11,		"")

// Keep the transaction date&time for reversal P90 field
DECLARE_SHARED_MEM_ARRAY( IcbcTransDateTime,	16,		"")

// Keep the state of PRR
	DECLARE_SHARED_MEM_ARRAY(IcbcPRRState,	6,	"9999")

// Keep the state of PRJ
	DECLARE_SHARED_MEM_ARRAY(IcbcPRJState,	6,	"9999")

// Keep the state of IDC
	DECLARE_SHARED_MEM_ARRAY(IcbcIDCState,	6,	"9999")

// Keep the state of CashIn module
	DECLARE_SHARED_MEM_ARRAY(IcbcCIMState,	6,	"9999")

// Keep the state of Cash Dispenser module
	DECLARE_SHARED_MEM_ARRAY(IcbcCDMState,	6,	"9999")

// Keep the cassette1 state of Cash dispenser 
	DECLARE_SHARED_MEM_ARRAY(IcbcCasState1,	6,	"9999")

// Keep the cassette2 state of Cash dispenser 
	DECLARE_SHARED_MEM_ARRAY(IcbcCasState2,	6,	"9999")

// Keep the cassette3 state of Cash dispenser 
	DECLARE_SHARED_MEM_ARRAY(IcbcCasState3,	6,	"9999")

// Keep the cassette4 state of Cash dispenser 
	DECLARE_SHARED_MEM_ARRAY(IcbcCasState4,	6,	"9999")

// Keep the total number of cassettes 
	DECLARE_SHARED_MEM_INT( IcbcTotCasNum,	0 )

// The total number of successful withdrawals for extra bank.
	DECLARE_REGISTRY_INTEGER( IcbcTotExtraWthNum,		0 )

// The total amount of withdrawal for extra bank.
	DECLARE_REGISTRY_DOUBLE( IcbcTotExtraWthAmount,		0.0 )

// The total number of withdrawal reversal for extra bank.
	DECLARE_REGISTRY_INTEGER( IcbcTotExtraWthRevNum,		0 )

// The total amount of withdrawal reversal for extra bank.
	DECLARE_REGISTRY_DOUBLE( IcbcTotExtraWthRevAmount,		0.0 )

// The current date of detail file to write
	DECLARE_SHARED_MEM_ARRAY(IcbcCurDetailFile,		20,		"")

// The date of detail file which the host asked for upload
	DECLARE_SHARED_MEM_ARRAY(IcbcDateOfUpLoad,		20,		"")
    
// The track3 update message received from Host
	DECLARE_SHARED_MEM_ARRAY(IcbcTrackUpdate,		10,		"")

// The host returned transaction index number. Printed on the journal
	DECLARE_SHARED_MEM_ARRAY(IcbcHostIndexNum,		25,		"")

// The host returned service charge fee, Printed on the journal and receipt.
	DECLARE_SHARED_MEM_ARRAY(IcbcHostServCharge,	15,		"")

	DECLARE_WAITABLE_SHARED_MEM_ARRAY( IcbcCashInAvail,		2,	"Y" )
//the settlementdate recieved from host for printing on journal
	DECLARE_SHARED_MEM_ARRAY(IcbcSettlementDate,		6,		"")

//the commission charge recieved from host for printing on prr
	DECLARE_SHARED_MEM_ARRAY(Icbccommicharge,		11,		"")

	DECLARE_SHARED_MEM_ARRAY(IcbcHostTime,		12,		"")

	DECLARE_SHARED_MEM_ARRAY(IcbcHostSeq,		66,		"")

// Keep the state of EPP
	DECLARE_SHARED_MEM_ARRAY(IcbcEPPState,	6,	"9999")

// add for emv project

// The hardward and software encrypt
	DECLARE_SHARED_MEM_ARRAY(GBLEncrypType,	2,	"")

// The trides and signle des encrypt
	DECLARE_SHARED_MEM_ARRAY(GBLEncrypMode,	2,	"")

// for hardware load key
	DECLARE_WAITABLE_SHARED_MEM_ARRAY(GBLLoadKeyStatus,  2,	"")

	DECLARE_SHARED_MEM_ARRAY(GBLDAMStatus,  2,	"C")

// Keep the state of FEP
	DECLARE_SHARED_MEM_ARRAY(DeviceFEPState,	5,	"0")

// Keep the state of PRR
	DECLARE_SHARED_MEM_ARRAY(DevicePRRState,	5,	"0")

// Keep the state of PRJ
	DECLARE_SHARED_MEM_ARRAY(DevicePRJState,	5,	"0")

// Keep the state of BAT
	DECLARE_SHARED_MEM_ARRAY(DeviceBATState,	5,	"0")

// Keep the state of OPD
	DECLARE_SHARED_MEM_ARRAY(DeviceOPDState,	5,	"0")

// Keep the state of Ooperator key
	DECLARE_SHARED_MEM_ARRAY(DeviceOPKState,	5,	"0")

// Keep the state of IDC
	DECLARE_SHARED_MEM_ARRAY(DeviceIDCState,	5,	"0")

// Keep the state of ICC
	DECLARE_SHARED_MEM_ARRAY(DeviceICCState,	5,	"0")

// Keep the state of SAM CARD
	DECLARE_SHARED_MEM_ARRAY(DeviceSAMState,	2,	"0")

// Keep the state of ALARM
	DECLARE_SHARED_MEM_ARRAY(DeviceALRState,	5,	"0")

// Keep the state of DEP
	DECLARE_SHARED_MEM_ARRAY(DeviceDEPState,	5,	"0")

// Keep the state of CashIn module
	DECLARE_SHARED_MEM_ARRAY(DeviceCIMState,	5,	"0")

// Keep the state of EDM
	DECLARE_SHARED_MEM_ARRAY(DeviceEDMState,	5,	"0")

// Keep the state of Cash Dispenser module
	DECLARE_SHARED_MEM_ARRAY(DeviceCDMState,	5,	"0")

// Keep the status of reject box for BOC
	DECLARE_SHARED_MEM_ARRAY(RejectBoxSts,	3,	"0")

// Keep the status of box1 for BOC
	DECLARE_SHARED_MEM_ARRAY(CashBoxSts1,	3,	"*")

// Keep the status of box2 for BOC
	DECLARE_SHARED_MEM_ARRAY(CashBoxSts2,	3,	"*")

// Keep the status of box3 for BOC
	DECLARE_SHARED_MEM_ARRAY(CashBoxSts3,	3,	"*")

// Keep the status of box4 for BOC
	DECLARE_SHARED_MEM_ARRAY(CashBoxSts4,	3,	"*")

// Keep the Trans JuLianDays for BOC
	DECLARE_SHARED_MEM_ARRAY(TransJulianDays,	6,	"0")

//Keep the Input PIN for PinChange for BOC
	DECLARE_SHARED_MEM_ARRAY(ShufflePinBlock,		 18,	"" )

// Increasing for each communication message 
	DECLARE_REGISTRY_INTEGER(Line8583Format, 0);

// Line Send Status (used in SNAcommdll) Y - already sent 
	DECLARE_SHARED_MEM_ARRAY(GBLSendStatus,  2,	"N")

// for Print
    DECLARE_SHARED_MEM_ARRAY( FitCardMark, 		    5,	"")

//I-系统刚启动 N-交换完密钥  R-密钥不同步，需要重新换密钥
  DECLARE_WAITABLE_SHARED_MEM_ARRAY( ResetTransKey,  2,	"N")

  //Y-主机cutoff  N-主机未cutoff
  DECLARE_WAITABLE_SHARED_MEM_ARRAY( HostCutOffFlag,  2,	"N")
    
  DECLARE_SHARED_MEM_ARRAY( GBLCashinResult ,3,"")
   DECLARE_SHARED_MEM_ARRAY( GBLCWDResult ,3,"")
  DECLARE_SHARED_MEM_ARRAY( GBLKeepAccountFlag ,2,"")

  //是否使用嵌入式网络数字硬盘录象机 EVR(Embedded Net DVR) Y - 使用 N - 不使用
	DECLARE_SHARED_MEM_ARRAY(GBLEVRUse,  2,	"N")
#include "DataLinkPost.h"