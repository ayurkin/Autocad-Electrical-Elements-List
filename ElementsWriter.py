class GroupsWriterToAutocadPage(object):
    def __init__(self, file_name, page_lines_count, modelspace, template_name):
        self.file_name = file_name
        self.page_lines_count = page_lines_count
        self.modelspace = modelspace
        self.template_name = template_name
        self.__file = self.create_file(self.template_name, self.modelspace)
        self.__left_groups = None

    def create_file(self, template_name, modelspace):
        return modelspace.Documents.Add(template_name)

    def save_file(self):
        self.__file.SaveAs(self.file_name)
        file_path = self.__file.FullName
        self.__file.Close()
        return file_path

    def write_groups(self, groups):
        pass


class GroupsWriterToAutocadFiles(object):
    ELEMENTS_LIST_PAGE_NAME = "elements_list"
    FIRST_PAGE_TEMPLATE = "a4-1-PE_v2.dwt"
    OTHER_PAGE_TEMPLATE = "a4-2-PE_v2.dwt"
    FORMAT = "dwg"
    FIRST_PAGE_LINES_COUNT = 10
    OTHER_PAGE_LINES_COUNT = 10

    def __init__(self):
        self.elements_list_files = []

    def write_groups(self, groups, modelspace):
        first_page_file_name = "{name}_1.{ext}".format(name=self.ELEMENTS_LIST_PAGE_NAME, ext=self.FORMAT)
        first_page_writer = GroupsWriterToAutocadPage(first_page_file_name, self.OTHER_PAGE_LINES_COUNT, modelspace,
                                                      self.FIRST_PAGE_TEMPLATE)
        first_page_writer.write_groups(groups)


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

    def write_elements(self, elements, modelspace):
        elements_sorted = self.get_sorted_by_tag(elements)  # Sort elements by tag
        groups = self.get_groups(elements_sorted) # Create groups by the same catalog number
        self.elements_list_files = self.groups_writer(groups, modelspace)




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