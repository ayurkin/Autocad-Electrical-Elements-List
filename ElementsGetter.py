def com_object_itervalues(com_object):
    for i in range(com_object.Count):
        yield com_object[i]


WASTE_ELEMENTS_NAMES = ["acade_title", "WD_M"]


class EntityFromAutocadSheet(object):
    FAMILY_ADAPTER = {"RES": "RE"}

    def __init__(self, entity_from_autocad_sheet):
        self.entity_attributes = {attribute.TagString: attribute.TextString
                                  for attribute in com_object_itervalues(entity_from_autocad_sheet.GetAttributes())}
        self.tag = self.entity_attributes.get("TAG1")
        self.producer = self.entity_attributes.get("MFG")
        self.catalog_number = self.entity_attributes.get("CAT")
        self.family = self.FAMILY_ADAPTER.get(self.entity_attributes.get("FAMILY"),
                                              self.entity_attributes.get("FAMILY"))

    def __str__(self):
        return "tag: {tag}, producer: {producer}, catalog number: {catalog_number}, family: {family}".format(
            tag=self.tag, producer=self.producer, catalog_number=self.catalog_number, family=self.family)


class EntityExporterFromModelSpace(object):
    def __init__(self, modelspace):
        self.modelspace = modelspace
        self.__entities = None

    def load_entities(self):
        for entity_from_autocad_sheet in com_object_itervalues(self.modelspace):
            if self.is_entity_ok(entity_from_autocad_sheet):
                self.__entities.append(EntityFromAutocadSheet(entity_from_autocad_sheet))

    @property
    def entities(self):
        if self.__entities is None:
            self.__entities = []
            self.load_entities()
        return self.__entities

    def is_entity_ok(self, entity_from_autocad_sheet):
        if not hasattr(entity_from_autocad_sheet, "name"):
            return False
        if not hasattr(entity_from_autocad_sheet, "getAttributes"):
            return False
        if entity_from_autocad_sheet.layer != "SYMS":
            return False
        if entity_from_autocad_sheet.name.startswith("WD"):
            return False
        if entity_from_autocad_sheet.name in WASTE_ELEMENTS_NAMES:
            return False

        return True


class ElementsGetter(object):

    def __init__(self, modelspace):
        self.__elements = None
        self.entity_exporter = EntityExporterFromModelSpace(modelspace)

    def load_elements(self):
        entities = [{"tag": entity.tag, "producer": entity.producer, "catalog_number": entity.catalog_number,
                     "family": entity.family, "description": "no description"} for entity in
                    self.entity_exporter.entities if entity.tag is not None]
        entities_without_catalog_number = [{"tag": entity.tag, "producer": entity.producer,
                                            "catalog_number": entity.catalog_number,
                                            "family": entity.family, "description": "no description"} for entity in
                                           self.entity_exporter.entities
                                           if entity.catalog_number is None]
        # if entities_without_catalog_number:
        #     raise Exception("There are elements without a catalog number:", [entity["tag"]
        #                                                                      for entity in
        #                                                                      entities_without_catalog_number])
        return entities

    @property
    def elements(self):
        if self.__elements is None:
            self.__elements = self.load_elements()
        return self.__elements
