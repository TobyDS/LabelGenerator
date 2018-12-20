# Import libraries
import win32com.client as client
import win32api
import wmi
import schedule
import time

with open('LabelTemplate.txt', 'r') as myfile:
    template=myfile.read()

# Set network adapter name as txtNetInterface

# THIS IS A TEMPORARY NETWORK ADAPTOR AND MUST BE COMMENTED OUT
txtNetInterface = "Realtek 8812AU Wireless LAN 802.11ac USB NIC"

# THIS IS THE REAL NETWORK ADAPTOR AND SHOULD NOT BE COMMENTED
#txtNetInterface = "Realtek RTL8811AU Wireless LAN 802.11ac USB 2.0 Network Adapter"


currentMAC = None

# Select all network adapters
c = wmi.WMI()
com = client.Dispatch("WbemScripting.SWbemRefresher")
obj = client.GetObject("winmgmts:\\root\cimv2")
allNetworkInterfaces = com.AddEnum(obj, "Win32_PerfRawData_Tcpip_NetworkInterface").objectSet

def MACAddressGetter():
    for interface in c.Win32_NetworkAdapterConfiguration ():
        # Select the interface with desired adapter name
        if interface.Description == txtNetInterface:
            interfaceOfInterest = interface
            interestDescription = interfaceOfInterest.Description
            interestMAC = interfaceOfInterest.MACAddress
            #return MAC Address
            return interestMAC

def changeDetector():
    newMAC = MACAddressGetter()
    global currentMAC
    if newMAC != currentMAC:
        if newMAC != None:
            MACNoColons = newMAC.replace(':','')
            outputPartial = template.replace('<MACwithcolons>',newMAC)
            outputFull = outputPartial.replace('<MAC>',MACNoColons)
            f = open("Label.txt", "w")
            f.write(outputFull)
            print newMAC
        currentMAC = newMAC

    
schedule.every(2).seconds.do(changeDetector)

while True:
    schedule.run_pending()
    time.sleep(1)
