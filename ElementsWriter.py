class ElementsListWriter(object):
    def __init__(self, groups_writer):
        self.groups_witer = groups_writer()

    def get_sorted(self, elements):
        s = sorted(elements, key=lambda element: element['tag'])

    def write_elements(self, elements):
        elements = self.get_sorted(elements)
