from difflib import SequenceMatcher
import re

def similarity(a, b):
	return round((SequenceMatcher(None, a, b).ratio() * 100), 1)

def string_comparator(a,b):
	# cleaning strings, tokenizing them
	a_tokens = list(a.strip())
	b_tokens = list(b.strip())
	# first try, blind assignment of max and min lists
	max_list = b_tokens
	min_list = a_tokens
	# check for reality: swap max/min lists if needed
	if(len(a) > len(b)):
		max_list = a_tokens
		min_list = b_tokens
	# configuring the 'likeness' degree (in %s)
	likeness = 0.0 # in percents
	likeness_unit = 100.0 / len(max_list)

	for i in range(len(min_list)):
		if(max_list[i] == min_list[i]):
			likeness += likeness_unit
	print("Likeness: {}%".format(likeness))

def get_mrf_number(prefix, ugly_mrf):
	result = ""
	m1 = re.search('(?<=' + prefix + ')\w+', ugly_mrf)
	m2 = re.search('(?<=' + prefix + ')\d+', ugly_mrf)
	if(m1 != None):	
		a = m1.group(0)
		result = a.strip()
	if(m2 != None):
		a = m2.group(0)
		resul = a.strip()
	return result


###### TEST BELOW ###### 

string_comparator("YAABBCCY", "AABBCC")
print(similarity("YACCY", "XACCY"))
# print(get_mrf_number("NCOC MRF ", "NCOC MRF 1234"))