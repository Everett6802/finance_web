from datetime import datetime, timedelta
import common_definition as CMN_DEF


def get_datetime_obj_fromm_query_time_strig(query_time_strig):
	# import pdb; pdb.set_trace()
	datetime_obj = datetime.strptime(query_time_strig, CMN_DEF.QUERY_TIME_STRING_FORMAT)
	return datetime_obj
