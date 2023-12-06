import time
from pylogix import PLC
from emailer import *

def write_header(file):
    file.write("curr_time, WeldReworkId, WeldReworkReasonCode\n")

current_hour = time.localtime().tm_hour

tag_list = ['WeldReworkReasonCode', 'WeldReworkId', 'WeldTechLogged']
reason_codes = {
    0: "Nothing",
    1: "Weld of location",
    2: "Weld burn through",
    3: "Weld Gap Issues",
    4: "Wire feeding issues",
    5: "Gap Issues/weld Porosity",
    6: "Bad tip/consumables",
    7: "Change of weld parameters"
}


while True:
    file_name = ""
    t = time.localtime()
    curr_time = str(t.tm_hour) + ":" + str(t.tm_min) + ":" + str(t.tm_sec)
    hour = int(time.strftime("%H"))
    dir = makeDir()
    curr_time = str(t.tm_hour) + ":" + str(t.tm_min) + ":" + str(t.tm_sec)

    # Check if it's time to generate the file (every 8 hours starting at 7 am)
    if curr_time == "7:0:0" or curr_time == "15:0:0" or curr_time == "23:0:0":
        hour = int(time.strftime("%H"))
        # make a new file
        file_name = dir + '\\' + generateFilename(hour)
        print("File name: "+ file_name)
        with open(file_name, 'w', newline='') as csv_file:
            write_header(csv_file)

    # Read from the PLC
    with PLC() as comm:
        comm.IPAddress = '10.10.14.12'
        ret = comm.Read(tag_list)

        if ret[2]:  # Writes out if our WeldTechlogged is True
            # print("This is my file name: ", curr_time)
            file_name = dir + '\\' + generateFilename(hour)
            with open(file_name, 'a') as csv_file:
                csv_file.write(f"{curr_time}, {ret[1].Value}, {reason_codes[ret[0].Value]}\n")

        time.sleep(25)
