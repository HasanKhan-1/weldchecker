
import csv
import time
from pylogix import PLC



current_hour = time.localtime().tm_hour

tag_list = ['WeldReworkReasonCode', 'WeldReworkId', 'WeldTechLogged']
t = time.localtime()
year = t.tm_year
month = t.tm_mon
day = t.tm_mday
curr_time = str(t.tm_hour) +":"+str(t.tm_min)+ ":"+ str(t.tm_sec)
print("The current time is: "+ curr_time+ " The date is: "+ str(year) +"/"+str(month)+"/"+ str(day))

# Check if it's time to generate the file (every 8 hours starting at 7 am)

if current_hour == 7 or (current_hour - 7) % 8 == 0:     
  hour = int(time.strftime("%H"))
  # make a new file 


elif current_hour == 0:
    # Send files
#   server = initEmailServer()
#   if server:
#       sendEmail(server)
#       server.quit()
#       print("Files sent at", datetime.now())
#   else:
#       print("Email server connection failed")
    print("dosomething")



with PLC() as comm:
    comm.IPAddress = '10.10.14.12'
    ret = comm.Read(tag_list)

    # for testing 
    for r in ret:
        print(r.Value)

    comm.Close()

    if (ret[2]):

        write_data = [('WeldReworkReasonCode', ret[0]),
                    ('WeldReworkId', ret[1])]

        with open('31_log.csv', 'w') as csv_file: # need to change the name of the file 
            csv_file = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_file.writerows(write_data)
            time.sleep(1)
    else:
        pass



