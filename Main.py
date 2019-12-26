from AutocadConnection import get_autocad_com_obj
from ElementsGetter import ElementsGetter
from CatalogInfoGetter import CatalogInfoGetter, CatalogInfoGetterFromSQLServer
from ElementsWriter import ElementsListWriter
from Constants import SQL_SERVER_CONSTANT_SERVER, SQL_SERVER_CONSTANT_DATABASE, \
    SQL_SERVER_CONSTANT_PASSWORD, SQL_SERVER_CONSTANT_UID

autocad_app = get_autocad_com_obj()

if autocad_app:
    model_space = autocad_app.ActiveDocument.ModelSpace

    elements_getter = ElementsGetter(model_space=model_space)
    elements_without_desc = elements_getter.elements

    sql_info_getter = CatalogInfoGetterFromSQLServer(
        server=SQL_SERVER_CONSTANT_SERVER,
        database=SQL_SERVER_CONSTANT_DATABASE,
        uid=SQL_SERVER_CONSTANT_UID,
        password=SQL_SERVER_CONSTANT_PASSWORD
    )

    catalog_info_getter = CatalogInfoGetter(
        elements_without_desc=elements_without_desc,
        info_getter_type=sql_info_getter
    )

    elements = catalog_info_getter.elements

    writer = ElementsListWriter()
    writer.write_elements(elements=elements, autocad_app=autocad_app)
