/******************************************************************************
*                                                                             *
*  DlHtmlWork.cpp: Defines Work fields in HTML screen                         *
*                                                                             *
******************************************************************************/

#include "DataLinkPre.h"

/////////////////////////////////////////////////////////////////////////////
	DECLARE_SHARED_MEM_ARRAY(HtmlWork11,		4,	"01" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork12,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork13,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork14,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork15,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork16,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork17,		10,	"" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWork21,		4,	"02" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork22,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork23,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork24,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork25,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork26,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork27,		10,	"" )
	
	DECLARE_SHARED_MEM_ARRAY(HtmlWork31,		4,	"03" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork32,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork33,		70,	"")
	DECLARE_SHARED_MEM_ARRAY(HtmlWork34,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork35,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork36,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork37,		15,	"" )
	
	DECLARE_SHARED_MEM_ARRAY(HtmlWork41,		4,	"04" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork42,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork43,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork44,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork45,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork46,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork47,		15,	"" )
	
	DECLARE_SHARED_MEM_ARRAY(HtmlWork51,		4,	"05" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork52,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork53,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork54,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork55,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork56,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork57,		15,	"" )
	
	DECLARE_SHARED_MEM_ARRAY(HtmlWork61,		4,	"06" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork62,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork63,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork64,		15,	"" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWork71,		4,	"07" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork72,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork73,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork74,		15,	"" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWork81,		4,	"08" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork82,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork83,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork84,		15,	"" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWork91,		4,	"08" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork92,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork93,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork94,		15,	"" )

	DECLARE_SHARED_MEM_ARRAY(HtmlWork101,		4,	"08" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork102,		15,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork103,		70,	"" )
	DECLARE_SHARED_MEM_ARRAY(HtmlWork104,		15,	"" )


/////////////////////////////////////////////////////////////////////////////
// Html Screen Title
DECLARE_SHARED_MEM_ARRAY( HtmlTitle,		40,	"")

// Html First input field1
DECLARE_SHARED_MEM_ARRAY( HtmlInput1,		24,	"")

// Html Second input field2
DECLARE_SHARED_MEM_ARRAY( HtmlInput2,		24,	"")

// Html First input field3
DECLARE_SHARED_MEM_ARRAY( HtmlInput3,		20,	"")

// Html Second input field4
DECLARE_SHARED_MEM_ARRAY( HtmlInput4,		20,	"")

// Html Screen Prompt1
DECLARE_SHARED_MEM_ARRAY( HtmlPrompt1,		40,	"")

// Html Screen Prompt2
DECLARE_SHARED_MEM_ARRAY( HtmlPrompt2,		40,	"")

// Html Screen Prompt3
DECLARE_SHARED_MEM_ARRAY( HtmlPrompt3,		40,	"")

// Html Screen Prompt4
DECLARE_SHARED_MEM_ARRAY( HtmlPrompt4,		40,	"")
	
// The balance of basic account returned 
DECLARE_SHARED_MEM_ARRAY(HtmlBasicBalance,		20,	"")

// for fashcash amount confirm 
DECLARE_SHARED_MEM_ARRAY( HtmlFastcashAmount,	12,	"" )

//for abc saving card inquire 
DECLARE_SHARED_MEM_ARRAY(HtmlUncheckNum,		5,	"")

// for fashcash amount confirm 
DECLARE_SHARED_MEM_ARRAY( HtmlAvailibleBalance,	20,	"" )
#include "DataLinkPost.h"