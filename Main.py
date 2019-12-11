from AutocadConnection import get_autocad_com_obj
from ElementsGetter import ElementsGetter
from CatalogInfoGetter import CatalogInfoGetter, CatalogInfoGetterFromSQLServer


SQL_SERVER_CONSTANT_SERVER = "TESTPC"
SQL_SERVER_CONSTANT_DATABASE = "AutocadNIO1"
SQL_SERVER_CONSTANT_UID = "sa"
SQL_SERVER_CONSTANT_PASSWORD = "niiset"

app = get_autocad_com_obj()
modelspace = app.ActiveDocument.ModelSpace



sql_info_getter = CatalogInfoGetterFromSQLServer(SQL_SERVER_CONSTANT_SERVER, SQL_SERVER_CONSTANT_DATABASE,
                                                 SQL_SERVER_CONSTANT_UID, SQL_SERVER_CONSTANT_PASSWORD)

elements_without_desc = ElementsGetter(modelspace).elements
elements = CatalogInfoGetter(elements_without_desc, sql_info_getter).elements



for element in elements:
    print element


