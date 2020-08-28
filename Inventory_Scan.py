import netmiko
from netmiko import ConnectHandler
import pandas as pd
import re
import xlsxwriter
import xlrd
import openpyxl
import getpass
from openpyxl import load_workbook


username = input("Enter Username: ")
password = getpass.getpass()

print ("\nScan Started!")

IP_Test = open('Location_IP.txt')

for ip in IP_Test:
    print ("\nStarting scan for ", ip )
    router = {
        'device_type': 'cisco_ios',
        'ip': ip,
        'username': username,
        'password': password
    }

    net_connect = ConnectHandler(**router)

    output_h = net_connect.send_command('show run | include hostname')

# Output for Cellular0
    searchfor_h = "hostname"
    res_h = output_h.find(searchfor_h)
# Check if the find value is valid
    if (output_h.find(searchfor_h) != -1):
        res_h1 = res_h + 9
        res_h2 = res_h + 15
        IPCELL_H = output_h[res_h1:res_h2]
        print(IPCELL_H)
    else:
        IPCELL_H = "unassigned"
        print(IPCELL_H)

    output_1 = net_connect.send_command('show ip int br')


#Output for Cellular0
    searchfor_c = "Cellular0"
    res_cell = output_1.find(searchfor_c)
#Check if the find value is valid
    if (output_1.find(searchfor_c) != -1):
        res_cell1 = res_cell + 27
        res_cell2 = res_cell + 38
        IPCELL_C = output_1[res_cell1:res_cell2]
        print (IPCELL_C)
    else:
        IPCELL_C = "unassigned"
        print(IPCELL_C)

#Output for Dialer0
    searchfor_d = "Dialer0"
    res_dialer = output_1.find(searchfor_d)
#Check if the find value is valid
    if (output_1.find(searchfor_d) != -1):
        res_dialer1 = res_dialer + 27
        res_dialer2 = res_dialer + 38
        IPCELL_D = output_1[res_dialer1:res_dialer2]
        print(IPCELL_D)
    else:
        IPCELL_D = "unassigned"
        print(IPCELL_D)

#Output for GigabitEthernet8
    searchfor_ge = "GigabitEthernet8"
    res_ge = output_1.find(searchfor_ge)
#Check if the find value is valid
    if (output_1.find(searchfor_ge) != -1):
        res_ge1 = res_ge + 27
        res_ge2 = res_ge + 38
        IPCELL_GE = output_1[res_ge1:res_ge2]
        print(IPCELL_GE)
    else:
        IPCELL_GE = "unassigned"
        print(IPCELL_GE)


    output_2 = net_connect.send_command('show cellular 0 all')

#Output for IMEI
    searchfor_imei = "(IMEI)"
    res_imei = output_2.find(searchfor_imei)
#Check if the find value is valid
    if (output_2.find(searchfor_imei) != -1):
        res_imei1 = res_imei + 9
        res_imei2 = res_imei + 24
        IPCELL_IMEI1 = output_2[res_imei1:res_imei2]
        IPCELL_IMEI = 0
        try:
            IPCELL_IMEI = IPCELL_IMEI + int(IPCELL_IMEI1)
            print(IPCELL_IMEI)

        except:
            IPCELL_IMEI = "unassigned"
            print(IPCELL_IMEI)
    else:
        IPCELL_IMEI = "unassigned"
        print(IPCELL_IMEI)

#Output for ICCID
    searchfor_iccid = "(ICCID)"
    res_iccid = output_2.find(searchfor_iccid)
#Check if the find value is valid
    if (output_2.find(searchfor_iccid) != -1):
        res_iccid1 = res_iccid + 10
        res_iccid2 = res_iccid + 29
        IPCELL_ICCID1 = output_2[res_iccid1:res_iccid2]
        IPCELL_ICCID = 0
        try:
            IPCELL_ICCID = IPCELL_ICCID + int(IPCELL_ICCID1)
            print(IPCELL_ICCID)

        except:
            IPCELL_ICCID = "unassigned"
            print(IPCELL_ICCID)
    else:
        IPCELL_ICCID = "unassigned"
        print(IPCELL_ICCID)

#Output for MSISDN
    searchfor_msisdn = "(MSISDN)"
    res_msisdn = output_2.find(searchfor_msisdn)
#Check if the find value is valid
    if (output_2.find(searchfor_msisdn) != -1):
        res_msisdn1 = res_msisdn + 11
        res_msisdn2 = res_msisdn + 22
        IPCELL_MSISDN1 = output_2[res_msisdn1:res_msisdn2]
        IPCELL_MSISDN = 0
        try:
            IPCELL_MSISDN = IPCELL_MSISDN + int(IPCELL_MSISDN1)
            print(IPCELL_MSISDN)

        except:
            IPCELL_MSISDN = "unassigned"
            print(IPCELL_MSISDN)

    else:
        IPCELL_MSISDN = "unassigned"
        print(IPCELL_MSISDN)

#Storing output to Excel

    df = pd.DataFrame ({"Location" : [str(IPCELL_H)],
                        "IP Address" : [str(ip)],
                        "Cellular 0" : [str(IPCELL_C)],
                        "Dialer 0" : [str(IPCELL_D)],
                        "GigabitEthernet0" : [str(IPCELL_GE)],
                        "IMEI" : [str(IPCELL_IMEI)],
                        "SIM #" : [str(IPCELL_ICCID)],
                        "Phone #" : [str(IPCELL_MSISDN)]
                        })
    writer = pd.ExcelWriter('Inventory_Excel.xlsx', engine = 'openpyxl')
    writer.book = load_workbook('Inventory_Excel.xlsx')
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    reader = pd.read_excel(r'Inventory_Excel.xlsx')
    df.to_excel (writer, sheet_name="Sheet1", index=False, header=False, startrow=len(reader)+1)

    writer.save()


print ("\nScan Ended!")