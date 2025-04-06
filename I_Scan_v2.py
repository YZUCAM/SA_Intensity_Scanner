"""Python code for Intensity scan
    By Dr. Yi Zhu in Cambridge Feb 2023
    This version is CLI based program without GUI
"""

import numpy as np
import thorlabs_apt as apt
import win32com.client
import win32gui
import time
import traceback

# key input parameters
path_directory = 'C:\\Users\Bay2 Users\Desktop\I_Scan_Data'
file_name = 'test1'
delay_time = 2  # unit second
# 1/2 plate rotation steps and total rotated range
angle_step = 2
angle_range = 30
initial_motor = 20


# available functions
def motor_move_to(motor, value):
    motor.move_to(value)


def motor_home(motor):
    motor.move_home()


def motor_clockwise(motor, step):
    motor.move_by(step)


def motor_anticlock(motor, step):
    motor.move_by(0-step)

#TODO modify this part of code to fit for this data structure
def save_data(path_name, xy):
    with open(path_name, "w") as file:
        for line in xy:
            v1, v2, v3 = line
            file.write('{0},  {1},  {2}\n'.format(v1, v2, v3))
        file.close()
        return "saved"


def I_scan_measure(motor, path_directory, angle_step, total_point, OphirCOM, DeviceHandle):
    time.sleep(5)
    print('start measurement: \n')
    row = list()
    dump1 = OphirCOM.GetData(DeviceHandle, 0)
    time.sleep(0.2)
    dump2 = OphirCOM.GetData(DeviceHandle, 1)
    time.sleep(0.2)
    for i in range(total_point):
        col = list()
        p1 = motor.position
        print('motor position: {0}'.format(int(p1)))
        data1 = OphirCOM.GetData(DeviceHandle, 0)
        time.sleep(0.2)
        data2 = OphirCOM.GetData(DeviceHandle, 1)
        time.sleep(0.2)
        col.append(p1)
        if len(data1[0]) > 0:
            print('sensor1: {0}'.format(data1[0][0]))
            col.append(data1[0][0])
        else:
            col.append(0)
        if len(data2[0]) > 0:
            print('sensor2: {0}'.format(data2[0][0]))
            col.append(data2[0][0])
        else:
            col.append(0)

        motor_clockwise(motor, angle_step)
        row.append(col)
        time.sleep(delay_time)
    #return row
    path_name = path_directory + '\\' + file_name + '.txt'
    save_data(path_name, row)

# initiate motor
motor = apt.Motor(83844170)
print('motor connected. \n')
motor_move_to(motor, initial_motor)

# initiate power meter
try:
    OphirCOM = win32com.client.Dispatch("OphirLMMeasurement.CoLMMeasurement")
    # Stop & Close all devices
    OphirCOM.StopAllStreams()
    OphirCOM.CloseAll()
    # Scan for connected Devices
    DeviceList = OphirCOM.ScanUSB()
    print(DeviceList)
    Device = DeviceList[0]
    #for Device in DeviceList:  # if any device is connected
    DeviceHandle = OphirCOM.OpenUSBDevice(Device)  # open first device
    exists1 = OphirCOM.IsSensorExists(DeviceHandle, 0)
    exists2 = OphirCOM.IsSensorExists(DeviceHandle, 1)
    if exists1 & exists2:
        print('\n----------Data for S/N {0} ---------------'.format(Device))

        sensor_info1 = OphirCOM.GetSensorInfo(DeviceHandle, 0)
        sensor_info2 = OphirCOM.GetSensorInfo(DeviceHandle, 1)
        print(sensor_info1)
        print(sensor_info2)
        # get sensor 1 range
        ranges1 = OphirCOM.GetRanges(DeviceHandle, 0)
        print(ranges1)
        ranges2 = OphirCOM.GetRanges(DeviceHandle, 1)
        print(ranges2)
            # start data streaming
        OphirCOM.StartStream(DeviceHandle, 0)
        OphirCOM.StartStream(DeviceHandle, 1)
    else:
        print('\nNo Sensor attached to {0} !!!'.format(Device))

    time.sleep(2)

    '''
    # main loop extract power
    while True:
        data1 = OphirCOM.GetData(DeviceHandle, 0)
        time.sleep(0.2)
        data2 = OphirCOM.GetData(DeviceHandle, 1)
        time.sleep(0.2)
        if len(data1[0]) > 0:
            print('sensor1: {0}'.format(data1[0][0]))
        if len(data2[0]) > 0:
            print('sensor2: {0}'.format(data2[0][0]))
        print('-----------')
        time.sleep(0.2)
    '''

    while True:
        print('########################')
        print('Choose your option?\n')
        print('1 Check sensor value.\n')
        print('2 Set path directory.\n')
        print('3 Set saved file name.\n')
        print('4 Set delay_time.\n')
        print('5 Set angle_step.\n')
        print('6 Set angle_range.\n')
        print('7 Set motor position.\n')
        print('8 Start I_scan measurement.\n')
        print('9 Check current path directory.\n')
        print('0 Exit.\n')

        choice = int(input('Input your option: \n'))
        if choice == 1:
            data1 = OphirCOM.GetData(DeviceHandle, 0)
            time.sleep(0.2)
            data2 = OphirCOM.GetData(DeviceHandle, 1)
            time.sleep(0.2)
            if len(data1[0]) > 0:
                print('sensor1: {0}'.format(data1[0][0]))
            if len(data2[0]) > 0:
                print('sensor2: {0}'.format(data2[0][0]))
            print('-----------')
            time.sleep(0.2)
        elif choice == 2:
            new_path_directory = input('input path directory: \n')
            path_directory = new_path_directory
        elif choice == 3:
            new_file_name = input('input new file name: \n')
            file_name = new_file_name
        elif choice == 4:
            new_delay_time = float(input('input delay time: \n'))
            delay_time = new_delay_time
        elif choice == 5:
            new_angle_step = float(input('input angle step: \n'))
            angle_step = new_angle_step
        elif choice == 6:
            new_angle_range = float(input('input angle range: \n'))
            angle_step = new_angle_range
        elif choice == 7:
            motor_set_pos = float(input('input motor initial position: \n'))
            initial_motor1 = motor_set_pos
            motor_move_to(motor, initial_motor1)
        elif choice == 8:
            total_point = int(angle_range/angle_step)
            data = I_scan_measure(motor, path_directory, angle_step, total_point, OphirCOM, DeviceHandle)
            print(data)
        elif choice == 9:
            print(path_directory)
        elif choice == 0:
            break
        else:
            pass

    '''
    total_point = int(angle_range/angle_step)
    data = I_scan_measure(motor, path_directory, angle_step, total_point, OphirCOM, DeviceHandle)
    print(data)
    '''


except OSError as err:
    print("OS error: {0}".format(err))
except:
    traceback.print_exc()











# Close motor
motor.close()
# Stop & Close all devices
OphirCOM.StopAllStreams()
OphirCOM.CloseAll()
# Release the object
OphirCOM = None
