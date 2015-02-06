'''
Code to change tree information
from xls file to json
'''

import xlrd
import numpy as np

# Functions
def column2list(data_column):
    data_list = []
    # Loop transforming unicode into ascii characters.
    for item in data_column:    
        value = item.value
        value_str = value.encode('ascii','ignore')
        data_list.append(value_str)
    return data_list

# Get data to xls file.
book = xlrd.open_workbook("Activity Net.xls")
anet_sheet = book.sheet_by_name('newNodes')
number_activities = anet_sheet.nrows

# Find major categories in sheet.
major_categories = anet_sheet.col_slice(3, 1, 'none')

# Change data type to get a list of string.
all_major_categories = column2list(major_categories)
all_major_categories2 = [x.value.encode('ascii', 'ignore') for x in major_categories]

# Removing repeated categories.
majors = list(set(all_major_categories))

# Create node Id to major categories.
node_id = range(1, len(majors)+1)

# Create structure in tree to major category.
results = {}
node_root = results[0] = {}
node_root['name'] = 'root'

for idx in node_id:
    node = results[idx] = {}
    node['name'] = majors[idx]
    if 'children' in node_root:
        node_root['children'].append(node)
    else:
        node_root['children'] = [node]

# Find second tier categories in sheet.
second_tier_categories = anet_sheet.col_slice(2, 1, 'none')

# Change data type to get a list of string
all_second_tier = column2list(second_tier_categories)

# Removing repeated categories.
second_tier = list(set(all_second_tier))

# Add nodes and children to tree structure for other categories.
counter = len(results)
for idx in range(1, len(all_second_tier)):
    new_node = all_second_tier[idx]    

    idx_new_node = all_second_tier.index(new_node)
    if idx_new_node < idx:
        continue
    parent_name = all_major_categories[idx]
    parent_id = majors.index(parent_name)
    
    node = results[counter] = {}
    node['name'] = new_node
    node_parent = results[parent_id+1] # add 1 because node 0 is root
    if 'children' in node_parent:
        node_parent['children'].append(node)
    else:
        node_parent['children'] = [node]
    counter += 1


# myarray = np.column_stack((col0, col1, col2, col3))

# out=json.dump(anetData,ensure_ascii=false)
