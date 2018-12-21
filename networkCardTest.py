# Import libraries
import win32com.client as client
import win32api
import wmi
import schedule
import time

# This opens the template for the lables and reads into a variable template
with open('LabelTemplate.txt', 'r') as myfile:
    template=myfile.read()

# Set network adapter name as txtNetInterface
# THIS IS A TEMPORARY NETWORK ADAPTOR AND MUST BE COMMENTED OUT
txtNetInterface = "Realtek 8812AU Wireless LAN 802.11ac USB NIC"

# THIS IS THE REAL NETWORK ADAPTOR AND SHOULD NOT BE COMMENTED
#txtNetInterface = "Realtek RTL8811AU Wireless LAN 802.11ac USB 2.0 Network Adapter"

# Set current MAC address as None for later use
currentMAC = None

# Select all network adapters
c = wmi.WMI()
com = client.Dispatch("WbemScripting.SWbemRefresher")
obj = client.GetObject("winmgmts:\\root\cimv2")
allNetworkInterfaces = com.AddEnum(obj, "Win32_PerfRawData_Tcpip_NetworkInterface").objectSet

# This function cheacks for the desired network adapter from the list of all network adapters
def MACAddressGetter():
    for interface in c.Win32_NetworkAdapterConfiguration ():
        # Select the interface with desired adapter name
        if interface.Description == txtNetInterface:
            interfaceOfInterest = interface
            # interestDescription = interfaceOfInterest.Description
            interestMAC = interfaceOfInterest.MACAddress
            #return MAC Address
            return interestMAC

# This function detects the MAC address of the network adapter when pugged in and creates a label
def changeDetector():
    # Uses MACAdressGetter function to get the MAC address of the network adapter
    # If the adapter is not enumerated then None is returned
    newMAC = MACAddressGetter()
    global currentMAC
    # Runs if the state has been changed
    if newMAC != currentMAC:
        # We are only interested in when the state has been changed to something that is not None
        if newMAC != None:
            # Uses replace feature to replace strings in the template
            MACNoColons = newMAC.replace(':','')
            outputPartial = template.replace('<MACwithcolons>',newMAC)
            outputFull = outputPartial.replace('<MAC>',MACNoColons)
            # A file called label is created and the contese of the lable is written to it
            f = open("Label.txt", "w")
            f.write(outputFull)
            # Alert to programme user
            print 'Network adapter detected'
            print 'MAC Address: '+ newMAC
            print 'Lable has been generated\n'
        # Contense of currentMAC must be updated whether it is None or a MAC address
        currentMAC = newMAC

# Sets up a timing interval and which function to run  
schedule.every(2).seconds.do(changeDetector)

# Runs the timing schedule that was created
while True:
    schedule.run_pending()
    time.sleep(1)
