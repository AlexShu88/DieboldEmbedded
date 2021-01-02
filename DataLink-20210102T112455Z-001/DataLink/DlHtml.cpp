//
// Always INCLUDE this file in your build !!
//
#include "DataLinkPre.h"

	DECLARE_VARIABLE		(HtmlFkeyMap, int, 4095)
	STORAGE_SHARED_MEM		(HtmlFkeyMap, int)

	NEEDS_FORMATTING		(HtmlFkeyMap, int)
	DEFAULT_FORMATTING		(HtmlFkeyMap, int)

	DECLARE_SHARED_MEM_ARRAY(HtmlFkeyList,		100, "")
	DECLARE_SHARED_MEM_ARRAY(HtmlSubstData,		128, "")
	DECLARE_SHARED_MEM_ARRAY(HtmlEditReturn,	 20, "")

	DECLARE_SHARED_MEM_ARRAY(HtmlScreenName,	128, "")


	DECLARE_VARIABLE		(MaintHtmlFkeyMap, int, 4095)
	STORAGE_SHARED_MEM		(MaintHtmlFkeyMap, int)

	NEEDS_FORMATTING		(MaintHtmlFkeyMap, int)
	DEFAULT_FORMATTING		(MaintHtmlFkeyMap, int)

	DECLARE_SHARED_MEM_ARRAY(MaintHtmlFkeyList,		100, "")
	DECLARE_SHARED_MEM_ARRAY(MaintHtmlSubstData,	128, "")
	DECLARE_SHARED_MEM_ARRAY(MaintHtmlEditReturn,	 20, "")

	DECLARE_SHARED_MEM_ARRAY(MaintHtmlScreenName,	128, "")

#include "DataLinkPost.h"
