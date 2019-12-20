from AutocadConnection import get_autocad_com_obj


#add file to project
#app.ActiveDocument.PostCommand("""(c:ace_add_dwg_to_project nil (list "" "" "" 1))\n""")

class GroupsWriterToAutocadPage(object):
    def __init__(self, file_name, page_lines_count, autocad_app, template_name):
        self.file_name = file_name
        self.page_lines_count = page_lines_count
        self.autocad_app = autocad_app
        self.template_name = template_name
        self.__file = self.create_file(self.template_name, self.autocad_app)
        self.__left_groups = None

    def create_file(self, template_name, autocad_app):
        return autocad_app.Documents.Add(template_name)

    def save_file(self):
        project_path = self.autocad_app.Documents[0].Path
        save_line = r"{project_path}\{file_name}".format(project_path=project_path, file_name=self.file_name)
        self.__file.SaveAs(save_line)
        # Execute command to add created file to project
        self.__file.SendCommand("""(c:ace_wdp_reread)\n""")
        self.__file.SendCommand("""(c:ace_add_dwg_to_project nil (list "" "" "" 1))\n""")
        # file_path = self.__file.FullName
        self.__file.Close()
        return None

    def write_groups(self, groups):
        groups = groups
        autocad_app = get_autocad_com_obj()
        current_file_obj = self.__file
        current_list_modelspace = self.__file.Modelspace
        table = None
        for i in range(current_list_modelspace.Count):
            if current_list_modelspace[i].EntityName == "AcDbTable":
                table = current_list_modelspace[i]

        table.SetCellValue(0, 0, "ABCDFJ55")
        for i in range(29):
            table.SetCellValue(i, 1, " ABCDFJ5513ABCDFJ5513ABCDFJ5513ABCDFJ5513ABCDFJ5513ABC")
            table.SetCellValue(i, 2, i)
        # table.SetCellValue(0, 1, u'\u041a\u043b\u0435\u043c\u043c\u0430 279-901 1.5')


        return []


class GroupsWriterToAutocadFiles(object):
    ELEMENTS_LIST_PAGE_NAME = "elements_list"
    FIRST_PAGE_TEMPLATE = "a4-1-PE_v2.dwt"
    OTHER_PAGE_TEMPLATE = "a4-2-PE_v2.dwt"
    FORMAT = "dwg"
    FIRST_PAGE_LINES_COUNT = 10
    OTHER_PAGE_LINES_COUNT = 10

    def __init__(self):
        self.elements_list_files = []

    def write_groups(self, groups, autocad_app):
        first_page_file_name = "{name}_1.{ext}".format(name=self.ELEMENTS_LIST_PAGE_NAME, ext=self.FORMAT)
        first_page_writer = GroupsWriterToAutocadPage(first_page_file_name, self.OTHER_PAGE_LINES_COUNT, autocad_app,
                                                      self.FIRST_PAGE_TEMPLATE)
        first_page_writer.write_groups(groups)
        file_path = first_page_writer.save_file()
        return []


class ElementsListWriter(object):
    GROUPS_WRITER = GroupsWriterToAutocadFiles

    def __init__(self):
        self.groups_writer = self.GROUPS_WRITER()
        self.elements_list_files = []

    def get_sorted_by_tag(self, elements):
        s = sorted(elements, key=lambda elem: elem['tag'])
        groups = [[s.pop(0)]]
        for element in s:
            prev_lit = get_literals_from_tag(groups[-1][0]['tag'])
            if prev_lit == get_literals_from_tag(element['tag']):
                groups[-1].append(element)
            else:
                groups.append([element])
        groups = [sorted(group, key=lambda el: get_digits_from_tag(el['tag'])) for group in groups]
        return sum(groups, [])

    def get_groups(self, elements):
        groups = [ElementsGroup([elements.pop(0)])]
        for element in elements:
            if groups[-1].can_be_added(element):
                groups[-1].add_element(element)
            else:
                groups.append(ElementsGroup([element]))
        return groups

    def write_elements(self, elements, autocad_app):
        elements_sorted = self.get_sorted_by_tag(elements)  # Sort elements by tag
        groups = self.get_groups(elements_sorted) # Create groups by the same catalog number
        self.elements_list_files = self.groups_writer.write_groups(groups, autocad_app)




class ElementsGroup(object):
    def __init__(self, elements):
        self.elements = elements

    def can_be_added(self, element):
        last_element = self.elements[-1]
        if element['catalog_number'] != last_element['catalog_number']:
            return False
        if get_digits_from_tag(element['tag']) - get_digits_from_tag(last_element['tag']) != 1:
            return False
        return True

    def add_element(self, element):
        self.elements.append(element)


def get_literals_from_tag(s):
    return "".join([d for d in s if not d.isdigit()])


def get_digits_from_tag(s):
    return int("".join([d for d in s if d.isdigit()]))


a = [{'description': "Kontaktor silovoi trehpolusniu 32A", 'tag': 'KM1', 'producer': "KEAZ", 'family': 'KM', 'catalog_number': '3.3'},
 {'description': "Kontaktor silovoi trehpolusniu 32A", 'producer': "KEAZ", 'family': 'KM', 'catalog_number': '3.3'},
 {'description': "Kontaktor silovoi trehpolusniu 32A", 'tag': 'KM3', 'producer': "KEAZ", 'family': 'KM', 'catalog_number': '3.3'}]

if len(a) == 3:
    fields = [{
        "tag": a[0]["tag"] + "-" + a[-1]["tag"],
        "desc": a[0]["description"] + " " + a[0]["producer"]

    }]



print fields

str_big = "Kontaktor silovoi trehpolusniu ofigennii big kosyakov i electromagnithih pomegh 32A"
print len(str_big)
print  str_big



while fields[-1]["desc"] > 30:

    fields.append({
        "tag": "",
        "desc": fields[-1]["desc"]
    })





    lst = str_big.split()
    str_big_ost = lst.pop(-1)
    str_big = " ".join(lst)
