# Use of Ophir COM object.
# Works with python 3.5.1 & 2.7.11
# Uses pywin32
import win32gui
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
    for Device in DeviceList:  # if any device is connected
        DeviceHandle = OphirCOM.OpenUSBDevice(Device)  # open first device
        exists1 = OphirCOM.IsSensorExists(DeviceHandle, 0)
        exists2 = OphirCOM.IsSensorExists(DeviceHandle, 0)
        if exists1 & exists2:
            print('\n----------Data for S/N {0} ---------------'.format(Device))

            # An Example for Range control. first get the ranges
            ranges = OphirCOM.GetRanges(DeviceHandle, 0)
            print(ranges)
            """
            # change range at your will
            if ranges[0] > 0:
                newRange = ranges[0] - 1
            else:
                newRange = ranges[0] + 1
            # set new range
            print('---------')
            print(newRange)
            print('---------\n')
            OphirCOM.SetRange(DeviceHandle, 0, newRange)
            """
            # An Example for data retrieving
            OphirCOM.StartStream(DeviceHandle, 0)  # start measuring
            OphirCOM.StartStream(DeviceHandle, 1)
            for i in range(10):
                time.sleep(.5)  # wait a little for data
                data = OphirCOM.GetData(DeviceHandle, 0)
                data2 = OphirCOM.GetData(DeviceHandle, 1)
                if len(data[0]) > 0:  # if any data available, print the first one from the batch
                    print('Reading1 = {0}, TimeStamp = {1}, Status = {2} '.format(data[0][0], data[1][0], data[2][0]))
                    #print('Reading2 = {0}, TimeStamp = {1}, Status = {2} '.format(data2[0][0], data2[1][0], data2[2][0]))
                    print(data)
                    print(data2)
                    print('---------------------------------------------------------------------------------------------')
        else:
            print('\nNo Sensor attached to {0} !!!'.format(Device))
except OSError as err:
    print("OS error: {0}".format(err))
except:
    traceback.print_exc()

win32gui.MessageBox(0, 'finished', '', 0)
# Stop & Close all devices
OphirCOM.StopAllStreams()
OphirCOM.CloseAll()
# Release the object
OphirCOM = None
