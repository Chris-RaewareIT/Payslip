# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 16:56:14 2022

@author: chris
"""

def unlock_PDF(src_File,trg_File):
    
    # import PdfFileWriter and PdfFileReader 
    # class from PyPDF2 library
    from PyPDF3 import PdfFileWriter, PdfFileReader
      
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
        for idx in range(file.numPages):
            
            # Get the page at index idx
            page = file.getPage(idx)
              
            # Add it to the output file
            out.addPage(page)
          
        # Open a new file "myfile_decrypted.pdf"
        with open(trg_File, "wb") as f:
            
            # Write our decrypted PDF to this file
            out.write(f)
      
        # Print success message when Done
        print("File decrypted Successfully.")
    else:
        
        # If file is not encrypted, print the 
        # message
        print("File already decrypted.")