# -*- coding: utf-8 -*-
"""
Created on Sat Nov 26 11:20:43 2022

@author: chris
"""
import pandas as pd
def main():

    ListObjectsInRootFolder()
    unpack_pdf()
    
def test_pdf_extract():
    strFullPath = 'C:/Users/chris/OneDrive/RITS/DCC/PaySlips/Unlocked/Week2 2022.pdf'
    
    strFileNoext ="Week2 2022"
    
    pdf_extract(strFullPath,strFileNoext)
    
def unpack_pdf():
    strRootFolder = 'C:/Users/chris/OneDrive/RITS/DCC/PaySlips/Unlocked/'
    
    lstObjects = getFolderobjects(strRootFolder)
    
    #lst_PayslipDetails =[]
    
    lst_PSD_TaxPeriod =[]
    lst_PSD_PeriodEnding =[]
    lst_PSD_EarningsForNICs =[]
    lst_PSD_NetPayment =[]
    lst_PSD_Pension_Employer =[]
    lst_PSD_Pension_Employee =[]
    lst_PSD_PensionAVC =[]


    
    for srcFile in lstObjects:
        file_name = srcFile.replace(strRootFolder,"")
        file_name_NoExt = file_name.replace(".pdf","")
        
        
        var_TaxPeriod,var_PeriodEnding,var_EarningsForNICs,var_NetPayment,var_Pension_Employer,var_Pension_Employee,var_PensionAVC = pdf_extract(srcFile,file_name_NoExt)
        
        
        #print(var_TaxPeriod)
        
        #print(type(var_PeriodEnding))
        
        
        lst_PSD_TaxPeriod.append(var_TaxPeriod)
        lst_PSD_PeriodEnding.append(var_PeriodEnding)
        
        lst_PSD_EarningsForNICs.append(var_EarningsForNICs)
        lst_PSD_NetPayment.append(var_NetPayment)
        lst_PSD_Pension_Employer.append(var_Pension_Employer)
        lst_PSD_Pension_Employee.append(var_Pension_Employee)
        lst_PSD_PensionAVC.append(var_PensionAVC)
        
        
    
    # print(lst_PSD_TaxPeriod)
    # print(lst_PSD_PeriodEnding)
    
    df_PayslipDetails = pd.DataFrame([lst_PSD_TaxPeriod,lst_PSD_PeriodEnding,lst_PSD_EarningsForNICs,lst_PSD_NetPayment,lst_PSD_Pension_Employer,lst_PSD_Pension_Employee,lst_PSD_PensionAVC]).T
    df_PayslipDetails.columns = ['TaxPeriod','PeriodEnding','EarningsForNICs','NetPayment','Pension_Employer','Pension_Employee','Pension_AVC']
    
    #print(df_PayslipDetails)
    
    strcsvFolder = "C:/Users/chris/OneDrive/RITS/DCC/PaySlips/csv/"
    
    csvPaySlipDetails = strcsvFolder + 'PaySlipDetails.csv'
    
    df_PayslipDetails.to_csv(csvPaySlipDetails, index=True, index_label="RowNumber", sep='|', encoding='utf-8', quotechar='"')

    
def pdf_extract(strInputFile,strFileNameNoext):
    import tabula
  


#strInputFileyFile = r"C:\Users\chris\OneDrive\RITS\DCC\PaySlips\Unlocked_NASA Umbrella Ltd Payroll for Chris Rae Week33 2022.pdf"

#strInputFileyFile = r"C:\Users\chris\OneDrive\RITS\DCC\PaySlips\NASA Umbrella Ltd Payroll for Chris Rae Week33 2022.pdf"

    #strInputFile = "myfile_decrypted.pdf"
    
    tables = tabula.read_pdf(strInputFile)
    
    df_01 = tables[0]
    
    # print(df_01)
    
    df_02 = tables[1]
  
    
    # print(df_02)
    
    
   # df_row_PensionEmp = df_01.loc[df_02['Unnamed: 0'] == "Employer's Pension"]
    df_row_PensionEmp = df_01[df_01['Unnamed: 0'].str.contains('Pension', na=False)]
    
    try:
    
        var_Pension_Employer = df_row_PensionEmp[df_row_PensionEmp.columns[1]].loc[df_row_PensionEmp.index[0]]
    except:
        var_Pension_Employer = 0
    
    # print(df_row_PensionEmp)
    
    # print (var_Pension_Employer)
    
    ############ Table 2
    
    df_row_TaxPeriod = df_02[df_02['Unnamed: 0'].str.contains('Tax Period', na=False)]
    
    
    
    var_TaxPeriod = df_row_TaxPeriod['Unnamed: 1'].loc[df_row_TaxPeriod.index[0]]
    
    
    # print (var_TaxPeriod)
    df_row_PeriodEnding = df_02[df_02['Unnamed: 0'].str.contains('Period Ending', na=False)]
    
    str_row_PeriodEnding =  df_row_PeriodEnding['Unnamed: 0'].loc[df_row_PeriodEnding.index[0]]
   
    
    var_PeriodEnding = str_row_PeriodEnding[-10:]
    
    
    df_row_EarningsForNICs = df_02[df_02['Unnamed: 2'].str.contains('Earnings for NICs', na=False)]
    
    var_EarningsForNICs = df_row_EarningsForNICs['PAYSLIP'].loc[df_row_EarningsForNICs.index[0]]
    
    
    
    df_row_NetPayment = df_02[df_02['Unnamed: 2'].str.contains('Net Payment', na=False)]
    
    var_NetPayment = df_row_NetPayment['PAYSLIP'].loc[df_row_NetPayment.index[0]]
    
    
    
    df_row_PensionEmployee = df_02[df_02['PAYSLIP'].str.contains('Pension Deductions', na=False)]
    
    try:
        var_Pension_Employee = df_row_PensionEmployee['Unnamed: 6'].loc[df_row_PensionEmployee.index[0]]
    except:
        var_Pension_Employee = 0
    
    
    df_row_PensionAVC = df_02.loc[df_02['PAYSLIP'] == 'Pension AVC']
    
    try:
        var_PensionAVC =  df_row_PensionAVC['Unnamed: 6'].loc[df_row_PensionAVC.index[0]]
    
    except:
        var_PensionAVC = 0
    
   

    
    

    
    
    return var_TaxPeriod,var_PeriodEnding,var_EarningsForNICs,var_NetPayment,var_Pension_Employer,var_Pension_Employee,var_PensionAVC
    
    
    
    
    

    
    
 
    
    
def unlock_PDF(src_File,trg_File,doneFile,strProblem):
    
    # import PdfFileWriter and PdfFileReader 
    # class from PyPDF2 library
    from PyPDF3 import PdfFileWriter, PdfFileReader
    import shutil
      
    # Create a PdfFileWriter object
    out = PdfFileWriter()
    strInputFileyFile = src_File
      
    # Open encrypted PDF file with the PdfFileReader
    file = PdfFileReader(strInputFileyFile)
    
      
    # Store correct password in a variable password.
    password = "20011971"
      
    # Check if the opened file is actually Encrypted
    if file.isEncrypted:
        
      
        # If encrypted, decrypt it with the password
        file.decrypt(password)
        
      
        # Now, the file has been unlocked.
        # Iterate through every page of the file
        # and add it to our new file.
        try:
            for idx in range(file.numPages):
                #print("Page numbers")
                
                # Get the page at index idx
                page = file.getPage(idx)
                  
                # Add it to the output file
                out.addPage(page)
              
            # Open a new file "myfile_decrypted.pdf"
            with open(trg_File, "wb") as f:
                
                # Write our decrypted PDF to this file
                out.write(f)
                
                
                shutil.move(src_File, doneFile)
            # Print success message when Done
            print("File decrypted Successfully.")
        except:
            shutil.move(src_File, strProblem)
    else:
        
        # If file is not encrypted, print the 
        # message
        print("File already decrypted.")
        shutil.move(src_File, trg_File)
    
def getFolderobjects(parDirectory):  
    import os
    dirObjects = os.listdir(parDirectory)
    dirObjectsfullpaths = map(lambda name: os.path.join(parDirectory, name), dirObjects)
    
    
    #df_objects = pd.DataFrame (dirObjectsfullpaths, columns = ['FullPath'])
    
    
    #print(dirObjects)
    #print(dirObjectsfullpaths)
    
    return dirObjectsfullpaths

def ListObjectsInRootFolder():
    
    


    # Get the Folders and files in the root folder  
    strRootFolder = 'C:/Users/chris/OneDrive/RITS/DCC/PaySlips/Locked/'
    strTargetFolder = 'C:/Users/chris/OneDrive/RITS/DCC/PaySlips/Unlocked/'
    
    strDoneFolder = 'C:/Users/chris/OneDrive/RITS/DCC/PaySlips/Done/'
    strProblem = 'C:/Users/chris/OneDrive/RITS/DCC/PaySlips/Problem/'
    
    
    print(strRootFolder)
    lstObjects = getFolderobjects(strRootFolder)
    
    for srcFile in lstObjects:
        file_name = srcFile.replace(strRootFolder,"")
        file_name = file_name.replace('NASA Umbrella Ltd Payroll for Chris Rae, ','')
        
        trgFile = strTargetFolder + file_name
        doneFile = strDoneFolder + file_name
        problemFile = strProblem + file_name
        print(trgFile)
        
        unlock_PDF(srcFile,trgFile,doneFile,problemFile)
        
        
    
    # df_objects = pd.DataFrame (lstObjects, columns = ['FullPath'])
    
    # df_objects["file_name"] = df_objects["FullPath"].str.replace(strRootFolder,"")
    # #df_objects["file_name"] = "chris"
    
    # print(df_objects)
    

    # df_Fld_Root = pd.DataFrame (strRootFolder,  columns = ['FullPath'])
    # lst_Fld_Root = df_Fld_Root['FullPath'].to_list()
    
    # FolderCount = df_Fld_Root.shape[0]
    
    
    # #print(lst_Fld_Root)
    
    # for folder in lst_Fld_Root:
        
        
    #     # Write the log entry for the start of the process
        
    #     try:
    #         print(folder)
    #         lstObjects = getFolderobjects(folder)
    #         df_objects = pd.DataFrame (lstObjects, columns = ['FullPath'])
    #         print(df_objects)
            
    #     except:
    #         pass
        
if __name__ == "__main__":
    main()

