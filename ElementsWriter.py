class ElementsListWriter(object):
    # def __init__(self, groups_writer):
        # self.groups_witer = groups_writer()

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

    def write_elements(self, elements):
        elements_sorted = self.get_sorted_by_tag(elements)
        groups = self.get_groups(elements_sorted)


class GroupsWriterToFiles(object):
    pass


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