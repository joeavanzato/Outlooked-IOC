#Requirements - pywin32 (Tested with Python 3.6.6 32-Bit and Pywin32-32bit) - Issues on 64-bit due to 32-bit Outlook usage on Hosts


import win32com.client, win32, win32com, os, argparse, re, time, traceback, sys
#print(win32com.__gen_path__) #Testing Purposes

#MD5/SHA1/SHA256 - 0-9, a-f, A-F (Hex Representations..Usually)
#MD5 - 32
#SHA1 - 40
#SHA256 - 64
#global md5, sha1, sha256

#Regex patterns for MD5/SHA1/SHA256 Hashes, IP Pattern, URL/Domain Pattern
md5_hex_pattern = re.compile(r'\b[a-f0-9]{32}\b', re.IGNORECASE)
sha1_hex_pattern = re.compile(r'\b[a-f0-9]{40}\b', re.IGNORECASE)
sha256_hex_pattern = re.compile(r'\b[a-f0-9]{64}\b', re.IGNORECASE)
#ip_pattern = re.compile(r'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b', re.IGNORECASE)
#url_pattern = re.compile(r'(http(s)?:\/\/.)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)', re.IGNORECASE)
ip_pattern = re.compile('(?:[\d]{1,3})\.(?:[\d]{1,3})\.(?:[\d]{1,3})\.(?:[\d]{1,3})', re.IGNORECASE)
url_pattern = re.compile('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', re.IGNORECASE)

#ArgParse Setting up cmd-line input - Utilize '-s' to specify folder structure to message feed -encapsulate in quotations if spaces exist
#If you utilize '/' in your Outlook folder naming conventions, this won't work.
parser = argparse.ArgumentParser(usage = '\n *-ISAC Message Parser and Hash Query\n  -s -- [Folder Structure storing FS-ISAC Messages]')
parser.add_argument("-s", "--structure", help='Specify structure to FS-ISAC message folder (I.E user@domain.com/Inbox/FS-ISAC OR user@domain.com/Confidential, etc', required=True)
args = parser.parse_args()
folder_structure = args.structure.split("/") #Split user input to determine folder for scanning
length_structure = len(folder_structure)
print("Total Items = "+str(length_structure))
outlook = win32com.client.Dispatch("Outlook.Application") #Required
mapi = outlook.GetNameSpace("MAPI") #Required

#This class is the entire program - designed to read Outlook mails sequentially when given specific mail folder
class mail_reader():
    global md5, sha1, sha256, ip, md5_list, sha1_list, sha256_list, ip_list
    def __init__(self):
        self.md5 = 0
        self.sha1 = 0
        self.sha256 = 0
        self.ip = 0
        self.URL = 0
        self.md5_list = []
        self.sha1_list = []
        self.sha256_list = []
        self.ip_list = []
        self.URL_list = []
        self.tmp = []

    def setup_mail_structure(self): #don't bury your fs-isac folder because I couldn't think of a better way to approach this and string concatenation withe mapi stuff is weird
        global mail_string
        for item in folder_structure:
            print(item)
        if length_structure == 2:
            mail_string = mapi.Folders[folder_structure[0]].Folders[folder_structure[1]]
        elif length_structure == 3:
            mail_string = mapi.Folders[folder_structure[0]].Folders[folder_structure[1]].Folders[folder_structure[2]]
        elif length_structure == 4:
            mail_string = mapi.Folders[folder_structure[0]].Folders[folder_structure[1]].Folders[folder_structure[2]].Folders[folder_structure[3]]
        elif length_structure == 5:
            mail_string = mapi.Folders[folder_structure[0]].Folders[folder_structure[1]].Folders[folder_structure[2]].Folders[folder_structure[3]].Folders[folder_structure[4]]
        elif length_structure == 6:
            mail_string = mapi.Folders[folder_structure[0]].Folders[folder_structure[1]].Folders[folder_structure[2]].Folders[folder_structure[3]].Folders[folder_structure[4]].Folders[folder_structure[5]]

#This function checks the body of an email and uses regex to find all instances of hashes, IPs or potential URLs/Domains, appending them to appropriate lists and setting flags based on detection
    def data_check(self, body):
        global md5, sha1, sha256, ip, url, md5_list, sha1_list, sha256_list, ip_list, URL_list
        self.md5_list = []
        self.sha1_list = []
        self.sha256_list = []
        self.ip_list = []
        self.URL_list = []
        try:
            self.tmp = []
            self.tmp = re.findall(md5_hex_pattern, body)
            for hash in self.tmp:
                self.md5_list.append(hash)
            if len(self.md5_list) != 0:
                self.md5 = 1
        except:
            self.md5 = 0
            #self.md5_list.append("N/A")
            #print(traceback.print_exc(sys.exc_info()))
            pass
        try:
            self.tmp = []
            self.tmp = re.findall(sha1_hex_pattern, body)
            for hash in self.tmp:
                self.sha1_list.append(hash)
            if len(self.sha1_list) != 0:
                self.sha1 = 1
        except:
            #self.sha1_list.append("N/A")
            self.sha1 = 0
            pass
        try:
            self.tmp = []
            self.tmp = re.findall(sha256_hex_pattern, body)
            for hash in self.tmp:
                self.sha256_list.append(hash)
            if len(self.sha256_list) != 0:
                self.sha256 = 1
        except:
            #self.sha256_list.append("N/A")
            self.sha256 = 0
        pass
        try:
            self.iplisttmp = []
            self.iplisttmp = re.findall(ip_pattern, body)
            #print(str(self.iplisttmp))
            for ip in self.iplisttmp:
                self.ip_list.append(ip)
            if len(self.ip_list) != 0:
                self.ip = 1
        except:
            #self.ip_list.append("N/A")
            self.ip = 0
            pass
        try:
            self.tmp = []
            self.tmp = re.findall(url_pattern, body)
            for url in self.tmp:
                self.URL_list.append(url)
            if len(self.URL_list) != 0:
                self.URL = 1
        except:
            #self.URL_list.append("N/A")
            self.URL = 0
        return self.md5, self.sha1, self.sha256, self.ip, self.URL

    #def write_hashes(self): #Write all hashes to single file with no other data
    #    with open("hash-list.txt", 'w') as f:


    #def write_IP(self): #Write all IP Addresses to single file with no other data
    #    pass

#This function iterates through all detected messages in a mailbox and performs file-writing to a CSV and TXT file to enable skipping of previously parsed messages and recording of detected data.
    def latest_email(self): #Read newest email, write data to 'email-lists.txt' file
        global md5, sha1, sha256, md5_list, sha1_list, sha256_list

        #self.skip = 0
        messages = mail_string.Items
        for email in list(messages):
            self.md5 = 0
            self.sha1 = 0
            self.sha256 = 0
            self.ip = 0
            self.URL = 0
            self.md5_list = []
            self.sha1_list = []
            self.sha256_list = []
            self.ip_list = []
            self.URL_list = []
            self.tmp = []
            self.skip = 0
            #email = messages.GetLast()
            body = email.body
            body = body.replace('hxxp', 'http')
            body = body.replace('[.]', '.')
            id = email.EntryID
            subject = email.Subject
            subject = subject.replace(",", "-")
            print(body)
            print(id)
            if os.path.isfile("email-lists.csv") == False: #Checks to see if file already exists due to opening with r+ in next stage, creates it otherwise
                with open("email-lists.csv", 'w') as f:
                    f.write("Message ID,"+"Message Subject,"+"IPs,"+"Domains,"+"MD5 Hashes,"+"SHA1 Hashes,"+"SHA256 Hashes"+"\n")
                    pass
            if os.path.isfile('ids+hashes.txt') == False:
                with open("ids+hashes.txt", 'w') as f:
                    f.write("MESSAGEID-IPs-Domains-MD5s-SHA1s-SHA256s"+"\n")
                    pass
            with open('ids+hashes.txt', 'r') as fe:
                for line in fe:
                    if id in line:
                        self.skip = 1
                        print("ID Detected in ids+hashes.txt - Already Scanned")
                        print("SKIPPED")
                        break
            if self.skip == 0:
                with open('email-lists.csv', 'a+') as f:
                    self.data_check(body)
                    f.write(id+","+subject+",")
                    if (self.md5+self.sha1+self.sha256+self.ip+self.URL) == 0:
                        print("Skipping - No Hashes/IPs/Domains Detected")
                        pass
                    else:
                        print("--- # IP Addresses / Domains Detected ---")
                        if self.ip == 1:
                            print("IP :" + str(len(self.ip_list)))
                            scan_list_ip = []
                            len_ip = len(self.ip_list)
                            print("\n" + "---IP ADDRESSES ---" + "\n")
                            for x in range(len_ip):
                                iptmp = self.ip_list[x]
                                f.write(str(iptmp) + " ")
                                scan_list_ip.append(iptmp)
                                print(iptmp)
                        else:
                            print("IP : 0")
                        f.write(",")
                        if self.URL == 1:
                            print("\n"+"URL :" + str(len(self.URL_list)))
                            scan_list_url = []
                            len_url = len(self.URL_list)
                            print("\n" + "---DOMAINS ---" + "\n")
                            for x in range(len_url):
                                urltmp = self.URL_list[x]
                                f.write(str(urltmp) + " ")
                                scan_list_url.append(urltmp)
                                print(urltmp)
                        else:
                            print("URL : 0")
                        f.write(",")
                        print("--- # Hashes Detected ---")
                        if self.md5 == 1:
                            print("\n"+"MD5 :" + str(len(self.md5_list)))
                            scan_list_md5 = []
                            len_md5 = len(self.md5_list)
                            print("\n" + "--- MD5 HASHES ---" + "\n")
                            for x in range(len_md5):
                                hash = self.md5_list[x]
                                f.write(hash + " ")
                                scan_list_md5.append(hash)
                                print(hash)
                        else:
                            print("MD5 : 0")
                        f.write(",")
                        if self.sha1 == 1:
                            print("\n"+"SHA1 :" + str(len(self.sha1_list)))
                            scan_list_sha1 = []
                            len_sha1 = len(self.sha1_list)
                            print("\n" + "--- SHA1 HASHES ---" + "\n")
                            for hash in self.sha1_list:
                                f.write(hash + " ")
                                print(hash)
                                scan_list_sha1.append(hash)
                        else:
                            print("SHA1 : 0")
                        f.write(",")
                        if self.sha256 == 1:
                            print("SHA256 :" + str(len(self.sha256_list)))
                            scan_list_sha256 = []
                            len_sha256 = len(self.sha256_list)
                            print("\n" + "--- SHA256 HASHES ---" + "\n")
                            for x in range(len_sha256):
                                hash = self.sha256_list[x]
                                f.write(hash + " ")
                                scan_list_sha256.append(hash)
                                print(hash)
                        else:
                            print("SHA256 : 0")
                    f.write("\n")
                with open('ids+hashes.txt', 'a+') as fe:
                    fe.write(id + "-" +str(self.md5_list)+ "-" + str(self.sha1_list) +"-"+ str(self.sha256_list)+"\n")
                        #f.write(id+ " Hashes_Detected-MD5: "+str(self.md5_list)+" SHA1: "+str(self.sha1_list)+" SHA256: "+str(self.sha256_list)+" IP: "+str(self.ip_list)+"\n")
            #self.write_hashes()
            #self.write_IP()


outlook = mail_reader() #Setting class
outlook.setup_mail_structure() #Setting MAPI/folder structure
#outlook.latest_email()

while True: #Infinite loop to check email-folder
    outlook.latest_email()
    time.sleep(5)
#read_mail_all()