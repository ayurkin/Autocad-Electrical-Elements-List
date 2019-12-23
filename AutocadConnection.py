import System
import sys


def get_autocad_com_obj():
    try:
        print "Try 'Autocad Electrical 2019'"
        autocad_com_obj = System.Runtime.InteropServices.Marshal.GetActiveObject(
            "Autocad.Application.23")  # Autocad 2019
    except Exception:
        print "'Autocad Electrical 2019' is not running"
        try:
            print "Try 'Autocad.Application.22'"
            autocad_com_obj = System.Runtime.InteropServices.Marshal.GetActiveObject("Autocad.Application")
        except Exception:
            print "'Autocad.Application.22' is not running"
            sys.exit()
    print "Autocad is running"
    return autocad_com_obj

