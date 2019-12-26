import System
from Constants import AUTOCADS


def get_autocad_com_obj():
    for autocad_name, autocad_connect_line in AUTOCADS.items():
        try:
            autocad_com_obj = System.Runtime.InteropServices.Marshal.GetActiveObject(autocad_connect_line)
        except EnvironmentError:
            pass
        else:
            print autocad_name, "is running"
            return autocad_com_obj
    else:
        print "AutoCAD is not running"
        return None
