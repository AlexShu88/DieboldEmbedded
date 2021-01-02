//
/////////////////////////////////////////////////////////////////////////////
//
//	DLGeneral.cpp
//
//	General shared variables
//
/////////////////////////////////////////////////////////////////////////////
//
#include "DataLinkPre.h"

#include <Time.h>
#include <stdio.h>

/////////////////////////////////////////////////////////////////////////////
	DECLARE_SHARED_MEM_ARRAY(MachineName,	21, "M U L T I M A T I C")
	DECLARE_SHARED_MEM_ARRAY(AgencyName,	27, "BACOB Trierstraat")
	DECLARE_SHARED_MEM_ARRAY(Version,		 6, "00.01")

/////////////////////////////////////////////////////////////////////////////
	ENCAPSULATE_ARRAY	(DateDDMMYY, char, 10, "")
	DECLARE_VARIABLE	(DateDDMMYY, ST_DateDDMMYY, stDateDDMMYYdefault)
	//	STORAGE_SHARED_MEM	(DateDDMMYY, ST_DateDDMMYY)
	short	SerializeDateDDMMYY(LPVOID lpBuffer, BOOL bRead)
	{
		if(bRead)
		{
			// Get the local Time
			struct tm	*newtime;
			time_t		 long_time;
			ST_DateDDMMYY	DateDDMMYY={0};
			int	tmp_Year;

			time (&long_time);
			newtime = localtime( &long_time ); /* Convert to local time. */

			tmp_Year = 1900 + newtime->tm_year;

			sprintf	(	(char *)&DateDDMMYY,
						"%02d/%02d/%04d",
						newtime->tm_mday,
						newtime->tm_mon + 1,
						tmp_Year
					);

			*(ST_DateDDMMYY *)lpBuffer=DateDDMMYY;
		}
		else
		{
			// Setting not supported
		}
		return (DL_SUCCESS);
	}
	//	STORAGE_SHARED_MEM	(DateDDMMYY, ST_DateDDMMYY)

	NEEDS_FORMATTING	(DateDDMMYY, ST_DateDDMMYY)
	DEFAULT_FORMATTING	(DateDDMMYY, char)
/////////////////////////////////////////////////////////////////////////////
	ENCAPSULATE_ARRAY	(TimeHHMMSS, char, 10, "")
	DECLARE_VARIABLE	(TimeHHMMSS, ST_TimeHHMMSS, stTimeHHMMSSdefault)
	//	STORAGE_SHARED_MEM	(TimeHHMMSS, ST_TimeHHMMSS)
	short	SerializeTimeHHMMSS(LPVOID lpBuffer, BOOL bRead)
	{
		if(bRead)
		{
			// Get the local Time
			struct tm	*newtime;
			time_t		 long_time;
			ST_TimeHHMMSS	TimeHHMMSS={0};

			time (&long_time);
			newtime = localtime( &long_time ); /* Convert to local time. */

			sprintf	(	(char *)&TimeHHMMSS,
						"%02d:%02d:%02d",
						newtime->tm_hour,
						newtime->tm_min,
						newtime->tm_sec
					);

			*(ST_TimeHHMMSS *)lpBuffer=TimeHHMMSS;
		}
		else
		{
			// Setting not supported
		}
		return (DL_SUCCESS);
	}
	//	STORAGE_SHARED_MEM	(TimeHHMMSS, ST_TimeHHMMSS)

	NEEDS_FORMATTING	(TimeHHMMSS, ST_TimeHHMMSS)
	DEFAULT_FORMATTING	(TimeHHMMSS, char)
/////////////////////////////////////////////////////////////////////////////

#include "DataLinkPost.h"
