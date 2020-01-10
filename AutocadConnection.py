import System
from Constants import AUTOCADS_CONNECTION_LINE


def get_autocad_com_obj():
    for autocad_name, autocad_connection_line in AUTOCADS_CONNECTION_LINE.items():
        try:
            autocad_com_obj = System.Runtime.InteropServices.Marshal.GetActiveObject(autocad_connection_line)
        except EnvironmentError:
            pass
        else:
            print autocad_name, "is running"
            return autocad_com_obj
    else:
        print "AutoCAD is not running"
        return None
