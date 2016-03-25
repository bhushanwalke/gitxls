__author__ = 'bhushan'

from pyexcel_xls import get_data
from pprint import pprint

data = get_data("CIS_CentOS_Linux_7_Benchmark_v1.1.0.xls")
pprint(data['Level 1'])