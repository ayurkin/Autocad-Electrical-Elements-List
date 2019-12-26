from time import sleep
from Constants import TABLE_ENTITY_NAME, TABLE_DESCRIPTION_LETTER_COUNT

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
        file_path = self.__file.FullName
        self.__file.Close()
        return file_path

    def write_groups(self, groups):
        left_lines_count = self.page_lines_count
        sleep(1)  # delay for 1 s to complete autocad file create
        current_list_modelspace = self.__file.Modelspace

        table = None
        for i in range(current_list_modelspace.Count):
            if current_list_modelspace[i].EntityName == TABLE_ENTITY_NAME:
                table = current_list_modelspace[i]

        while groups:
            group = groups.pop(0)
            lines = group.get_lines()
            if len(lines) >= left_lines_count:
                self.__left_groups = groups
                break
            line_number = self.page_lines_count - left_lines_count

            for line in lines:
                table.SetCellValue(line_number, 0, line["tag"])
                table.SetCellValue(line_number, 1, line["desc"])
                table.SetCellValue(line_number, 2, line["count"])
                line_number += 1

            left_lines_count -= len(lines)

    def get_left_groups(self):
        return self.__left_groups


class GroupsWriterToAutocadFiles(object):
    ELEMENTS_LIST_PAGE_NAME = "elements_list"
    FIRST_PAGE_TEMPLATE = "a4-1-PE_v2.dwt"
    OTHER_PAGE_TEMPLATE = "a4-2-PE_v2.dwt"
    FORMAT = "dwg"
    FIRST_PAGE_LINES_COUNT = 29
    OTHER_PAGE_LINES_COUNT = 322

    def __init__(self):
        self.elements_list_files = []

    def write_groups(self, groups, autocad_app):
        first_page_file_name = "{name}_1.{ext}".format(
            name=self.ELEMENTS_LIST_PAGE_NAME,
            ext=self.FORMAT
        )
        first_page_writer = GroupsWriterToAutocadPage(
            file_name=first_page_file_name,
            page_lines_count=self.FIRST_PAGE_LINES_COUNT,
            autocad_app=autocad_app,
            template_name=self.FIRST_PAGE_TEMPLATE
        )
        first_page_writer.write_groups(groups)
        file_path = first_page_writer.save_file()
        self.elements_list_files.append(file_path)
        left_groups = first_page_writer.get_left_groups()
        page_number = 2

        while left_groups:
            writer_page_file_name = "{name}_{page_number}.{ext}".format(
                name=self.ELEMENTS_LIST_PAGE_NAME,
                page_number=page_number,
                ext=self.FORMAT)

            writer = GroupsWriterToAutocadPage(writer_page_file_name,
                                               self.OTHER_PAGE_LINES_COUNT, autocad_app,
                                               self.OTHER_PAGE_TEMPLATE)
            page_number += 1
            writer.write_groups(left_groups)
            left_groups = writer.get_left_groups()
            file_path = writer.save_file()
            self.elements_list_files.append(file_path)

        return self.elements_list_files


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
        groups = self.get_groups(elements_sorted)  # Create groups by the same catalog number
        self.elements_list_files = self.groups_writer.write_groups(groups, autocad_app)
        print "Created files:"
        for file_name in self.elements_list_files:
            print file_name


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

    def word_iterator(self, words_list):
        for word in words_list:
            yield word

    def get_lines(self):
        lines = [{
            "tag": "",
            "desc": "",
            "count": ""
        }]
        first_element = self.elements[0]
        second_element = self.elements[1]
        last_element = self.elements[0]

        if len(self.elements) == 1:
            lines[0]["tag"] = first_element["tag"]
            lines[0]["count"] = str(1)
        elif len(self.elements) == 2:
            lines[0]["tag"] = first_element["tag"] + ", " + second_element["tag"]
            lines[0]["count"] = str(2)
        elif len(self.elements) >= 3:
            lines[0]["tag"] = first_element["tag"] + "-" + last_element["tag"]
            lines[0]["count"] = str(len(self.elements))

        desc = self.elements[0]["description"] + " " + self.elements[0]["producer"]
        desc_list = desc.split()

        for word in self.word_iterator(desc_list):
            if len(lines[-1]["desc"] + " " + word) > TABLE_DESCRIPTION_LETTER_COUNT:
                lines.append({"tag": "", "desc": " " + word, "count": ""})
            else:
                lines[-1]["desc"] = lines[-1]["desc"] + " " + word

        return lines


def get_literals_from_tag(s):
    return "".join([d for d in s if not d.isdigit()])


def get_digits_from_tag(s):
    return int("".join([d for d in s if d.isdigit()]))
