import re
from datetime import datetime, timedelta
# import common_definition as CMN_DEF
from common import common_definition as CMN_DEF


def get_datetime_obj_from_query_time_string(query_time_string, suppress_exception=True):
	# import pdb; pdb.set_trace()
	if re.match("[\d]{2}:[\d]{2}_[\d]{4}[\d]{2}[\d]{2}", query_time_string) is None:
		if suppress_exception:
			return None
		raise ValueError("Incorrect time string format: %s" % query_time_string)
	datetime_obj = datetime.strptime(query_time_string, CMN_DEF.API_DATETIME_STRING_FORMAT)
	return datetime_obj
