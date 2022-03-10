'''
Prerequisite: Install pandas respective to your OS.
DESC: This program reads the XML files and extracted Error,Message,Rule object from it and load it into ELSX file into it's respective sheets. This program also generate Log file for testing purposes.
Step1: Ask user input for Directory of XML files
Step2: Once the user enter the Directory the program will traverse through each of an XML file  and processed it only if it meets all the condition and pass through all the flags. 
Condition to extract the objects from XML:
            1) XML file should be in the directory
            2) The root of the XML file should be Data 
            3) It should match the path in order to fetch the data 
                   Errors=>  Data/Nuula/Errors/Error      
                   Message=> Data/DataExtract900jer/Messages/Message
                   Rules=> Data/DataExtract900jer/Rules/Rule
To run this program the above condition should be meet. If not, an appropriate error will generate on Log as well as Terminal.
In the program I declared 5 lists:
    A[] = Has the information of the Error object.
    B[] = Has the information of the Message object.
    C[] = Has the information of the Rules object.
    D[] = Used as a flag, to check the count. This tells when the root is not Data.
    E[] = Used as a flag, to check the count. This tells when their is XML file present in the provided directory.

Logic behind this: Extracted the data in the form of Dictionary from XML and put it into lists because I'm using Pandas DF and it take 2 dimensional objects. Once retrieved I used that List of Dictionaries of store the data into Excel. Once we get the data into their respective List then this program makes an Excel sheet with respective the information extracted.
'''
import os # For user input
import xml.etree.cElementTree as ET # To access the XML
import pandas as pd # To put the extracted data into XLSX file
import logging # To make a log file 

# Initializing and creatign object for logging to use in the program
logging.basicConfig(filename="Nuula.log",format='%(asctime)s, %(levelname)-8s [%(filename)s] %(message)s')
logger = logging.getLogger()#Creating an object
logger.setLevel(logging.INFO) #Setting the threshold of logger
A,B,C,D,E = [],[],[],[],[] # Made a list to store values and raise flag on conditions
incorrectPath = 0
rootNotData = 0
dir = input("Please enter home directory:")
for filename in os.listdir(dir):
    if not filename.endswith('.xml'): continue
    fullname = os.path.join(dir,filename)
    E.append(fullname)
    tree1 = ET.parse(fullname)
    root1 = tree1.getroot()
    if root1.tag == 'Data':
        D.append(fullname)  
        for error in root1.findall("./Nuula/Errors/Error"):
            A.append(error.attrib)
        for message in root1.findall("./DataExtract900jer/Messages/Message"):
            B.append(message.attrib)
        for rule in root1.findall("./DataExtract900jer/Rules/Rule"):
            C.append(rule.attrib)
        if len(E)>0 and len(A) ==0 and len(B)==0 and len(C)==0:
          incorrectPath+=1
          print("Incorrect path of file" + filename)
          logging.info("Incorrect path" + filename)  
    else:
        rootNotData+=1
        print("Root of an XML file " + filename + " is not Data")
        logging.info("Root of an XML file " + filename + " is not Data")
if len(E)==0 and len(A) ==0 and len(B)==0 and len(C)==0:
    print("No XML file found")    
    logging.info("No XML file found")
if len(E)>0 and len(A)>0 or len(B)> 0 or len(C)>0 :
      numProcessed = len(E)-(incorrectPath+rootNotData)
      dfError = pd.DataFrame(A)
      dfMessage = pd.DataFrame(B)
      dfRule = pd.DataFrame(C)
      writer = pd.ExcelWriter('Nuula.xlsx', engine='xlsxwriter')  
      if len(A) > 0 : dfError.to_excel(writer, sheet_name='Error') 
      if len(B) > 0 : dfMessage.to_excel(writer, sheet_name='Message') 
      if len(C) > 0 : dfRule.to_excel(writer, sheet_name='Rules')
      writer.save()
      print("Number of XML files in the directory are:",len(E)) #correct
      print("Number of XMLs processed:",numProcessed)
      logging.info("Number of files processed: " + str(numProcessed))
      logging.info('ELSX file created')
