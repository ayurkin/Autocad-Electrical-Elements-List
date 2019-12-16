from AutocadConnection import get_autocad_com_obj
from ElementsGetter import ElementsGetter
from CatalogInfoGetter import CatalogInfoGetter, CatalogInfoGetterFromSQLServer
import json

SQL_SERVER_CONSTANT_SERVER = "TESTPC"
SQL_SERVER_CONSTANT_DATABASE = "AutocadNIO1"
SQL_SERVER_CONSTANT_UID = "sa"
SQL_SERVER_CONSTANT_PASSWORD = "niiset"

app = get_autocad_com_obj()
modelspace = app.ActiveDocument.ModelSpace

# sql_info_getter = CatalogInfoGetterFromSQLServer(SQL_SERVER_CONSTANT_SERVER, SQL_SERVER_CONSTANT_DATABASE,
#                                                  SQL_SERVER_CONSTANT_UID, SQL_SERVER_CONSTANT_PASSWORD)

# elements_without_desc = ElementsGetter(modelspace).elements
# elements = CatalogInfoGetter(elements_without_desc, sql_info_getter).elements

# with open("data_file.json", "w") as f:
#     json.dump(elements, f, indent=4)

with open("data_file.json", "r") as f:
    elements = json.load(f)

for element in elements:
    print element


def get_literals_from_tag(s):
    return "".join([d for d in s if not d.isdigit()])


def get_digits_from_tag(s):
    return "".join([d for d in s if d.isdigit()])


def get_sorted(elements):
    s = sorted(elements, key=lambda elem: elem['tag'])
    groups = [[s.pop(0)]]
    for element in s:
        prev_lit = get_literals_from_tag(groups[-1][0]['tag'])
        if prev_lit == get_literals_from_tag(element['tag']):
            groups[-1].append(element)
        else:
            groups.append([element])

    # for group in groups:
    #     print "group start"
    #     print group
    #     print "group end"

    groups = [sorted(group, key=lambda elem: int(get_digits_from_tag(elem['tag']))) for group in groups]

    return sum(groups, [])


y = get_sorted(elements)

print "end"
