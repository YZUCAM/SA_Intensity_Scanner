'''test for ophir power meter'''


import win32com.client
import time
import traceback

try:
    OphirCOM = win32com.client.Dispatch("OphirLMMeasurement.CoLMMeasurement")
    # Stop & Close all devices
    OphirCOM.StopAllStreams()
    OphirCOM.CloseAll()
    # Scan for connected Devices
    DeviceList = OphirCOM.ScanUSB()
    print(DeviceList)

except OSError as err:
    print("OS error: {0}".format(err))
except:
    traceback.print_exc()


# Stop & Close all devices
OphirCOM.StopAllStreams()
OphirCOM.CloseAll()
# Release the object
OphirCOM = None