"""
Code to change tree information
from xls file to json
"""
# Functions.


# Add nodes and children to tree structure for other categories.
def add_node_to_structure(results_tmp, bias, all_new_nodes,
                          all_parent_categories, parents):
    counter = len(results_tmp)
    for idx in range(0, len(all_new_nodes)):
        new_node = all_new_nodes[idx]
        idx_new_node = all_new_nodes.index(new_node)
        if idx_new_node < idx:
            continue

        parent_name = all_parent_categories[idx]
        parent_id = parents.index(parent_name)

        node = results_tmp[counter] = {}
        node['name'] = new_node
        # add number of previous node.
        parent_id = parent_id+bias
        node_parent = results_tmp[parent_id]
        if 'children' in node_parent:
            node_parent['children'].append(node)
        else:
            node_parent['children'] = [node]
        counter += 1
    return results_tmp

# Get data to xls file.
import xlrd
book = xlrd.open_workbook("../Activity Net.xls")
anet_sheet = book.sheet_by_name('newNodes')
number_activities = anet_sheet.nrows

# Find major categories in sheet.
major_categories = anet_sheet.col_slice(3, 1, 'none')
# Change data type to get a list of string.
all_major_categories = [x.value.encode('ascii', 'ignore')
                        for x in major_categories]
# Removing repeated categories
majors = list(set(all_major_categories))
# Create structure in tree to major category.
results = {}
node_root = results[0] = {}
node_root['name'] = 'root'
# majors = []   # List of major categories not repeated.
count = 1
for index in majors:
    major_node = results[count] = {}
    major_node['name'] = index
    if 'children' in node_root:
        node_root['children'].append(major_node)
    else:
        node_root['children'] = [major_node]
    count += 1

# Find second tier categories in sheet.
second_tier_categories = anet_sheet.col_slice(2, 1, 'none')
# Change data type to get a list of string.
all_second_tier = [x.value.encode('ascii', 'ignore')
                   for x in second_tier_categories]
# Removing repeated categories
second_tier = list(set(all_second_tier))
# Add nodes of second tier category.
offset = 1   # previous node is only root
results = add_node_to_structure(results, offset, all_second_tier,
                                all_major_categories, majors)

# Find third tier category in sheet.
third_tier_categories = anet_sheet.col_slice(1, 1, 'none')
# Change data type to get a list of string.
all_third_tier = [x.value.encode('ascii', 'ignore')
                  for x in third_tier_categories]
# Removing repeated categories
third_tier = list(set(all_third_tier))
# Add nodes of third categories.
offset += len(majors)
results = add_node_to_structure(results, offset, all_third_tier,
                                all_second_tier, second_tier)

# Find activity names list in sheet.
activity_name_categories = anet_sheet.col_slice(0, 1, 'none')
# Change data type to get a list of string.
all_activity_name = [x.value.encode('ascii', 'ignore')
                     for x in activity_name_categories]
# Add nodes of third categories.
offset += len(second_tier)
results = add_node_to_structure(results, offset, all_activity_name,
                                all_third_tier, third_tier)
# Save structure in json file.
anetData = results[0]
import json
q = json.dumps(anetData, ensure_ascii=False)
with open('new_nodes.json', 'w') as outfile:
    outfile.write(unicode(q))
print q
