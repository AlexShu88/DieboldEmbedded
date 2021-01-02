#include <Windows.h>

#include "DataLinkGeneral.h"

// =============================================================================
//  = Formatting of an integer to a string and reverse                        =
//  = Should be replaced by PcFormat...                                       =
// =============================================================================
short Format_int(LPVOID Internal, LPVOID External, LPSTR Format, int nDirection, int size)
{
// PEC
	if (strcmpi (Format, "char") )
	{
		// Only format from and to char allowed
		return (DL_UNDEFINED);
	}
// end PEC

	switch(nDirection)
	{
		case FORMAT_INTERNAL_TO_SPECIFIED:
// PEC
//			itoa(**((int **)Internal), (LPSTR)External, 10);
			itoa(*((int *)Internal), (LPSTR)External, 10);
// end PEC
			return (DL_SUCCESS);
		break;

		case FORMAT_SPECIFIED_TO_INTERNAL:
// PEC
//			**(int **)Internal=atoi((LPSTR)External);
			*(int *)Internal=atoi((LPSTR)External);
// end PEC
			return (DL_SUCCESS);
		break;

		default:
			return (DL_UNDEFINED);
		break;
	}
}


// PEC
short Format_char(LPVOID Internal, LPVOID External, LPSTR Format, int nDirection, int size)
{
// PEC
	if (strcmpi (Format, "int") )
	{
		// Only format from and to int allowed
		return (DL_UNDEFINED);
	}

	switch(nDirection)
	{
		case FORMAT_INTERNAL_TO_SPECIFIED:
			*(int *)External=atoi((LPSTR)Internal);
			return (DL_SUCCESS);
		break;

		case FORMAT_SPECIFIED_TO_INTERNAL:
			itoa(*((int *)External), (LPSTR)Internal, 10);
			return (DL_SUCCESS);
		break;

		default:
			return (DL_UNDEFINED);
		break;
	}
}
// end PEC

