import clr
clr.AddReference("System.Data")
from System.Data import SqlClient, DataSet


class DBConnectionOpen:

    def __init__(self, connection):
        self.connection = connection

    def __enter__(self):
        self.connection.Open()
        return self.connection

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.connection.Close()
        return True


class CatalogInfoGetter(object):
    def __init__(self, elements_without_desc, info_getter_type):
        self.elements_without_desc = elements_without_desc
        self.catalog_info_getter = info_getter_type

    @property
    def elements(self):
        elements_desc = self.catalog_info_getter.get_elements_info(self.elements_without_desc)

        for element in self.elements_without_desc:
            element["description"] = elements_desc[element["catalog_number"]]

        return self.elements_without_desc


class CatalogInfoGetterFromSQLServer(object):

    def __init__(self, server, database, uid, password):
        self.server = server
        self.database = database
        self.uid = uid
        self.password = password
        self.connection = None
        self.connect_to_database()

    def connect_to_database(self):
        self.connection = SqlClient.SqlConnection(
            r"server={server};database={database};uid={uid};password={password}".format(
                server=self.server,
                database=self.database,
                uid=self.uid,
                password=self.password)
        )

    def get_queries(self, elements):
        queries = {}
        for element in elements:
            if element["family"] in queries:
                queries[element['family']].add(element["catalog_number"])
            else:
                queries[element["family"]] = {element["catalog_number"]}

        queries_strings = []
        for table, catalog_numbers in queries.iteritems():
            query_end = " OR ".join(
                [r"[CATALOG]={cat_num}".format(cat_num=cat_num) for cat_num in catalog_numbers]
            )

            query = r"""SELECT [CATALOG], [DESCRIPTION] FROM [{database}].[dbo].[{table}]
                        WHERE {query_end}""".format(
                database=self.database,
                table=table,
                query_end=query_end
            )
            queries_strings.append(query)
        return queries_strings

    def get_elements_info(self, elements):
        queries = self.get_queries(elements)
        catalog_info = {}
        for query in queries:
            adapter = SqlClient.SqlDataAdapter(query, self.connection)
            with DBConnectionOpen(self.connection):
                ds = DataSet()
                adapter.Fill(ds)
                row_indexies = {str(name): i for i, name in enumerate(ds.Tables[0].Columns)}

                for row in ds.Tables[0].Rows:
                    catalog_info[row[row_indexies["CATALOG"]]] = row[row_indexies["DESCRIPTION"]]

        return catalog_info
