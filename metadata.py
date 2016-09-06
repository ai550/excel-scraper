columns = {1:'A', 2:'B', 3:'C', 4:'D', 5:'E', 6:'F', 7:'G', 8:'H', 9:'I', 10:'J', 11:'K', 12:'L'}

mdata_dict = {'fl':'From Location', 'tl':'To Location', 'mrf':'Ref Number', 'date':'Date'}


# array that will hold 2 values, column and row number for each found label
# and in the end shall hold the value of that field as well

def get_metadata_labels(sheet, fl_label, tl_label, mrf_label, date_label):
	for row in range(1, sheet.max_row):
		for col in range(1, sheet.max_column):
			if(col in columns):
				c_value = str(sheet[columns[col] + str(row)].value).strip()
				# we have 4 pieces of information
				# first, let's label locations
				if(c_value == mdata_dict['fl'] or mdata_dict['fl'] in c_value):
					# print('[from location] location: ' + columns[col] + str(row))
					fl_label[0] = col
					fl_label[1] = row
				if(c_value == mdata_dict['tl'] or mdata_dict['tl'] in c_value):
					# print('[to location] location: ' + columns[col] + str(row))
					tl_label[0] = col
					tl_label[1] = row
				if(c_value == mdata_dict['mrf'] or mdata_dict['mrf'] in c_value):
					# print('[ref number] location: ' + columns[col] + str(row))
					mrf_label[0] = col
					mrf_label[1] = row
				if(c_value == mdata_dict['date'] or mdata_dict['date'] in c_value):
					# print('[date] location: ' + columns[col] + str(row))
					date_label[0] = col
					date_label[1] = row

# now that we found labels, let's search for their respective values

# function for finding values based on label and last column provided
def get_metadata_value(sheet, label, last_column):
	for col in range(label[0] + 1, last_column):
		if(col in columns):
			c_value = sheet[columns[col] + str(label[1])].value
		if(c_value != None):
			return c_value

