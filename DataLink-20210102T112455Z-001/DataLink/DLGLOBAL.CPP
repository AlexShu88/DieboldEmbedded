/******************************************************************************
*                                                                             *
*  DlGolbal.cpp: Defines application global data                              *
*                                                                             *
******************************************************************************/
#include "DataLinkPre.h"

#include <Time.h>
#include <stdio.h>
/********************************************************************/

/*********************************************************************
*																	 * 	
*			Define Global Config Field								 *
*																	 *
*********************************************************************/	


// application environment --> 0=off-line; 1=on-line; 2=on-off   
DECLARE_SHARED_MEM_ARRAY( GBLSysEnvir,		2,	"" )

//The Institution Code of Bank 
DECLARE_SHARED_MEM_ARRAY( GBLBankCode,		10,	"")

//For ICBC_HQ Begin modify by lijun
// Branch Code 
DECLARE_SHARED_MEM_ARRAY( GBLBranchCode,	10,	"")

// ATM Indentification Code
DECLARE_SHARED_MEM_ARRAY( GBLAtmCode,		10,	"")
//For ICBC_HQ End modify

// Sequence of reading card tracks --> 3=Track3; 2=Track2
DECLARE_SHARED_MEM_ARRAY( GBLTrackPriority,	4,	"")

// Capture card indicator : Y/N
DECLARE_SHARED_MEM_ARRAY( GBLCardCapture,	2,	"" )

// Transaction serial number for comm. 
// Do not need reset during open period
DECLARE_SHARED_MEM_ARRAY( GBLLineNum,		7,	"" )

// Contract Telephone Number
//Define from Global.ini
DECLARE_SHARED_MEM_ARRAY( GBLPhoneNumber,		12,	"" )

// Audio help control field
// Defined by S3EMonitor, used by HTML with audio embedded
// "Y" -- Play
// "N" -- Not play
DECLARE_SHARED_MEM_ARRAY( GBLAudioControl,		2,	"N" )

// Define current user selected language
// 'CHN' ---- Chinese;  'ENG' ----- English;
// 'OTH' ---- Others ;   default is 'CHN'
DECLARE_SHARED_MEM_ARRAY( GBLSelectLan,	4,	"CHN" )

// The max. notes can be delivered in one trans.
	DECLARE_SHARED_MEM_ARRAY(GBLMaxBills,	4,	"40")
/*******************************************************************/


/********************************************************************
*																	*
*			Define the ATM Runtime Status							*
*																	*
/*******************************************************************/
  // Operating Status corresponding to Operator Key Position 
// 1=Operator;	2=Customer;	3=Supervisor
DECLARE_SHARED_MEM_ARRAY( GBLOperStatus,	2,	"")

// External line status: O = Active; C = Down 
DECLARE_WAITABLE_SHARED_MEM_ARRAY( GBLLineStatus,	2,	"" )

// External line status: In idle = 'O'; In out of service = 'C' 
// Corresponding to the system situation, whether is in "idle" or 
// in "out of service"
DECLARE_SHARED_MEM_ARRAY( GBLAtmStatus,	2,	"C" )

// ATM Hardware status
// if there are any key atm modules failure, 
//		such as badger and journal, this feild is 'C';
// otherwise, the field is 'O'.
DECLARE_SHARED_MEM_ARRAY( GBLHWStatus,	2,	"C" )

// ATM period status
// Period Opened = 'O'; Period closed = 'C'
DECLARE_WAITABLE_REGISTRY_ARRAY( GBLPeriodStatus,	2,	"" )

// ATM Host management status
// ATM closed by host = 'C'; ATM opened by host = 'O'
// Normally, the status is 'O' until ATM receive a Close ATM command
// from host
// if the status is 'C', even reboot the field should remain as 'C'.
DECLARE_WAITABLE_REGISTRY_ARRAY( GBLHostCmdStatus,	2,	"" )

// The trigger to inform OutOfServ of doing LoadKey operation.
//It should be set in "S3EMonitor" and "OutOfServ" module before exit.
DECLARE_WAITABLE_SHARED_MEM_ARRAY( GBLHostDownloadStatus,	2,	"" )

// ATM host download key status
// ATM start download key ='C'; ATM finished the download key = 'O'
// The trigger to inform Monitor module of doing "DoRecovery" for 
//some device if necessary. 
//It should be set in "EndVisit" and "Operator" module before exit.
DECLARE_WAITABLE_SHARED_MEM_ARRAY( GBLDoRecovery,	2,	"" )

// The trigger to inform Monitor module of taking the Anomaly message 
// from Anomaly Queue
//It should be set in "EndVisit" and "Operator" module before exit.
DECLARE_WAITABLE_SHARED_MEM_ARRAY( GBLGetAnomalies, 2,	"" )

// Starter module status
// Starter module is running = "R"; Starter module has closed = "C"  
DECLARE_SHARED_MEM_ARRAY( GBLStarterStatus,	2,	"" )

// Defined two different types of receipt printer for CMB Shenzhen
// 'T' ---- for diebold standard thermal receipt printer
// 'I' ---- for IBM 4783 impact receipt printer or other T-957B impact receipt printer
DECLARE_SHARED_MEM_ARRAY( GBLRECPrinterType,4,	"" )


// The trigger to inform Monitor module of taking the initial State of Cassettes 
DECLARE_WAITABLE_SHARED_MEM_ARRAY( GBLInitCasStates,	2,	"" )

//  Device states for PACE STA2 Format
DECLARE_SHARED_MEM_ARRAY( GBLDevice_State,     66,  "" )

// AutoReboot status: P = Prepare; S = Ready; I = No Reboot 
DECLARE_WAITABLE_SHARED_MEM_ARRAY( GBLSysShutDown,	2,	"" )

/*******************************************************************/


/********************************************************************
*																	*
*		Define the software encryption key for Pin and MAC			*
*																	*
********************************************************************/

// Master key for software encryption, use to decrypt the following
// Work key and primary Mac key for individual generating the 
// Pinkey and MAC key
DECLARE_SHARED_MEM_ARRAY( GBLMasterKey,	50,	"" )

// Work key is the primary Pinkey
DECLARE_SHARED_MEM_ARRAY( GBLPrePinKey,	50,	"" )

// Primary Mac Key, 
DECLARE_SHARED_MEM_ARRAY( GBLPreMacKey,	50,	"" )

/********************************************************************
*																	*
*		Indicate the cassette information predefine in Global.ini   *
*																	*
********************************************************************/
// For multi currency withdraw
DECLARE_REGISTRY_ARRAY(GBLCashAvailCNY, 2, "Y")

// For multi currency withdraw
DECLARE_REGISTRY_ARRAY(GBLCashAvailHKD, 2, "N")

// Indicate whether it is a new period when loading cassette. (Y/N)
DECLARE_REGISTRY_ARRAY(GBLIsNewPeriod, 2, "")

// Indicate whether it is doing SDOCdm.Recovery. (Y/N)
DECLARE_SHARED_MEM_ARRAY(GBLIsDoRecoverying, 2, "")

// The denomination of Cassette box number 6
DECLARE_SHARED_MEM_INT( GBLCdmRecoveryTimes,	0 )

// A flag to indicate if the CDM module is needed recovery.(Y/N)
DECLARE_SHARED_MEM_ARRAY( GBLCdmRecoveryNeeded,	2, "Y" )


/*******************************************************************

/********************************************************************
*																	*
*		Define the general transaction field shared in different 
*		modules														*
*																	*
********************************************************************/
// The general transaction amount field, not seperator
DECLARE_SHARED_MEM_ARRAY( GBLAmount,	12,	"" )

// The general transaction amount field for printing on Prr and Prj, 
// including the "," and "."
DECLARE_SHARED_MEM_ARRAY( GBLPrtAmount,	16,	"" )

// The another transaction serial number field, getting it from 
// S3eLineout.GetData function after sending message OK.
// This field is used for printing on receipt and for sending 
// reversal. 
DECLARE_SHARED_MEM_ARRAY( GBLLineSendNum,	8,	"" )

// Transaction serial number received from host. 
DECLARE_SHARED_MEM_ARRAY( GBLRetLineNum,	7,	"" )

DECLARE_SHARED_MEM_ARRAY( GBLBRCMFlag, 2, "" )

// If the Pin has been validated by Host, the value is "Y";
// else the value is "N". 
// This value should be set to "N" in idle start.
DECLARE_SHARED_MEM_ARRAY( GBLIsPinOk,	2,	"" )

// Fit table download status. Used for Idle to fresh Fit table in memory
DECLARE_WAITABLE_REGISTRY_ARRAY( GBLFreshFitTable,	2,	"" )

// Define the transactions available status in Registry
// Inquiry	= 0x01
// PinChange = 0x02
// Transfer Out = 0x04
// Withdrawal = 0x08
// CashIn = 0x10
// Deposit = 0x20
// Transfer In = 0x40
// Agent = 0x80
DECLARE_REGISTRY_INTEGER( GBLTransAvailable,		0 )

// Define the select account type for HK JECTO card in misc field for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY( GBLAccType,	2,	"" )

// Define the two bits ATM local reject code for CMB Shenzhen
DECLARE_SHARED_MEM_ARRAY( GBLATMLocRejCode, 5, "")

//Define the PIN Mode
DECLARE_SHARED_MEM_ARRAY( GBLPinMode, 2, "S")

//Define the flag of 
DECLARE_REGISTRY_ARRAY( GBLIsKeyInitOK,	2,	"N" )

//Define the flag of Turn Off Audio Again  
DECLARE_SHARED_MEM_ARRAY( GBLAudioOffAgain,	2,	"N" )

/*** Indicate the current Date  **********************************/
	ENCAPSULATE_ARRAY	(GBLDateYYYYMMDD, char, 12, "")
	DECLARE_VARIABLE	(GBLDateYYYYMMDD, ST_GBLDateYYYYMMDD, stGBLDateYYYYMMDDdefault)
	//	STORAGE_SHARED_MEM	(DateDDMMYY, ST_DateDDMMYY)
	short	SerializeGBLDateYYYYMMDD(LPVOID lpBuffer, BOOL bRead)
	{
		if(bRead)
		{
			// Get the local Time
			struct tm	*newtime;
			time_t		 long_time;
			ST_GBLDateYYYYMMDD	GBLDateYYYYMMDD={0};
			int	tmp_Year;

			time (&long_time);
			newtime = localtime( &long_time ); /* Convert to local time. */

			tmp_Year = 1900 + newtime->tm_year;

			sprintf	(	(char *)&GBLDateYYYYMMDD,
						"%04d/%02d/%02d",
						tmp_Year,
						newtime->tm_mon + 1,
						newtime->tm_mday
					);

			*(ST_GBLDateYYYYMMDD *)lpBuffer=GBLDateYYYYMMDD;
		}
		else
		{
			// Setting not supported
		}
		return (DL_SUCCESS);
	}
	//	STORAGE_SHARED_MEM	(DateDDMMYY, ST_DateDDMMYY)

	NEEDS_FORMATTING	(GBLDateYYYYMMDD, ST_GBLDateYYYYMMDD)
	DEFAULT_FORMATTING	(GBLDateYYYYMMDD, char)
/****************************************************************/

/** Indicate the current time ***********************************/
	ENCAPSULATE_ARRAY	(GBLTimeHHMM, char, 7, "")
	DECLARE_VARIABLE	(GBLTimeHHMM, ST_GBLTimeHHMM, stGBLTimeHHMMdefault)
	//	STORAGE_SHARED_MEM	(GBLTimeHHMM, ST_GBLTimeHHMM)
	short	SerializeGBLTimeHHMM(LPVOID lpBuffer, BOOL bRead)
	{
		if(bRead)
		{
			// Get the local Time
			struct tm	*newtime;
			time_t		 long_time;
			ST_GBLTimeHHMM	GBLTimeHHMM={0};

			time (&long_time);
			newtime = localtime( &long_time ); /* Convert to local time. */

			sprintf	(	(char *)&GBLTimeHHMM,
						"%02d:%02d",
						newtime->tm_hour,
						newtime->tm_min
					);

			*(ST_GBLTimeHHMM *)lpBuffer=GBLTimeHHMM;
		}
		else
		{
			// Setting not supported
		}
		return (DL_SUCCESS);
	}
	//	STORAGE_SHARED_MEM	(GBLTimeHHMM, ST_GBLTimeHHMM)

	NEEDS_FORMATTING	(GBLTimeHHMM, ST_GBLTimeHHMM)
	DEFAULT_FORMATTING	(GBLTimeHHMM, char)
/****************************************************************/

#include "DataLinkPost.h"