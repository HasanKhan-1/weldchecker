# weldchecker @ MartinreaHFS

main.py generates a csv file containing the Weld Tally reports at 7AM, 3PM and 11PM, and emailer.py sends these three files at 7AM (?) every day to the mailing list. 

If the boolean WeldTechLogged is True, it adds the WeldReworkID and the associated Reason code to the csv.
