from openpyxl import load_workbook

path = 'C:\\Users\\nk91009578.NCOC\\Desktop\\'
wb = load_workbook(path + 'acmt_cupboard_1.xlsx')

cupboard1_ws = wb['Cupboard#1']
cupboard2_ws = wb['Cupboard#2']
cupboard3_ws = wb['Cupboard#3']

result_ws = wb['Sheet3']
my_dict = {} # dictionary, the value is iterated when duplicate found, thus the inventory

"""
Function takes a dictionary and an iterable.
Loop goes through iterable and populates dictionary.
If loop encounters a duplicate of any key in dictionary,
the value of that key is incremented.
"""
def populate_dict(dictionary, iterable):
    for row in iterable:
        key = str(row[0].value)
        value = row[2].value
        if key in dictionary:
            dictionary[key] = dictionary[key] + 1
        else:
            dictionary[key] = value

# populate_dict(my_dict, cupboard1_ws.rows)
# populate_dict(my_dict, cupboard2_ws.rows)
# populate_dict(my_dict, cupboard3_ws.rows)

def group
