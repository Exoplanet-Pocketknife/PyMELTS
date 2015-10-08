import os, sys, xlrd, xlwt, shutil, subprocess, time, pyautogui, inspect


#!/usr/bin/env python


#this is the first thing the program does, the central user interface from which everything else is run
def initialization():
    """

    :rtype :
    """
    print "Welcome to the Exoplanet Pocketknife.  Enter 'BSP,' 'Basalt,' 'Credits,' or 'Help' to begin"
    first_choice = raw_input(["Please enter 'BSP,' 'Basalt,' 'Credits,' or 'Help' to begin: "])
    if first_choice == "BSP":
        wholeplanet_writefiles()
        run_bsp()
    elif first_choice == "Basalt":
        basalt_writefiles()
        run_basalt()
    elif first_choice == "Credits":
        print "alphaMELTS algorithm by Ghiorso et. al \ Exoplanet Pocketknife by Scott D. Hull and Cayman T. Unterborn"
    elif first_choice == "Help":
        print "See documentation for further instructions..." #create documentation later...
    else:
        print "Oops!  That's not a valid command."
        initialization() #loop back to the beginning of the code



#this will write MELTS files used for the bsp calculation by using "whole planet" spectral data
def wholeplanet_writefiles():
    try:
        if not os.path.exists('Whole_Planet_MELTS_Files'):
            os.mkdir('Whole_Planet_MELTS_Files')
        else:
            print "Directory already exists.  Deleting and recreating..."
            shutil.rmtree('Whole_Planet_MELTS_Files')
            os.mkdir('Whole_Planet_MELTS_Files')

        os.path.join('Whole_Planet_MELTS_Files')

        xl_workbook = xlrd.open_workbook(raw_input("Please enter your workbook name: "), 'rb')
        sheet_names = xl_workbook.sheet_names()
        print ('Sheet Names', sheet_names)

        xl_sheet = xl_workbook.sheet_by_index(0)
        print ('Sheet name: %s' % xl_sheet.name)


        num_cols = xl_sheet.ncols #numbers of columns

        for j in range(xl_sheet.nrows):
            row = xl_sheet.row(j)
            #print row
            print "Writing MELTS file..."
            file_name = str(row[0].value)
            file_name_clipped = file_name[7:] #gets rid of the Title: for the filename
            melts_file = open('Whole_Planet_MELTS_Files/'+file_name_clipped.rstrip()+ '.txt', 'w') #rstrip gets rid of the whitespace at the end of the filename string
            for i in range(num_cols):
                melts_file.write(str(row[i].value)+'\n') #from what I've read you set the row in line 23, then just iterate over each column by value = row[column]
            melts_file.close() #close the file, move on to the next one

    except:
        print "Error!  Unable to write MELTS files." #Error message
        initialization() #loop back to the beginning of the code



#this will write basalt MELTS files after the BSP calculation and core subtraction
def basalt_writefiles():
    """

    :rtype :
    """
    try:
        if not os.path.exists('BSP_MELTS_Files'):
            os.mkdir('BSP_MELTS_Files')
        os.path.join('BSP_MELTS_Files')

        xl_workbook = xlrd.open_workbook(raw_input("Please enter your workbook name: "), 'r')
        sheet_names = xl_workbook.sheet_names()
        print ('Sheet Names', sheet_names)

        xl_sheet = xl_workbook.sheet_by_index(0)
        print ('Sheet name: %s' % xl_sheet.name)


        num_cols = xl_sheet.ncols #numbers of columns

        for j in range(xl_sheet.nrows):
            row = xl_sheet.row(j)
            #print row
            print "Writing MELTS file..."
            file_name = str(row[0].value)
            file_name_clipped = file_name[7:] #gets rid of the Title: for the filename
            melts_file = open('BSP_MELTS_Files/'+file_name_clipped.rstrip()+ '.txt', 'w') #rstrip gets rid of the whitespace at the end of the filename string
            for i in range(num_cols):
                melts_file.write(str(row[i].value)+'\n') #from what I've read you set the row in line 23, then just iterate over each column by value = row[column]
            melts_file.close() #close the file, move on to the next one

    except:
        print "Error!  Unable to write MELTS files." #Error message
        initialization() #loop back to the beginning of the code


#automates alphaMELTS for basalt calculations
def run_bsp():
    if not os.path.exists('Completed_Whole_Planet_MELTS_Files'):
        os.mkdir('Completed_Whole_Planet_MELTS_Files') #create a directory for successful output files
    os.path.join('Completed_Whole_Planet_MELTS_Files')

    current_file = os.path.basename("Whole_Planet_MELTS_Files") #for use when copying and renaming the alphaMELTS_tbl.txt file
    current_file_minus_txt = current_file[:-3]

    f = "run_alphamelts.command -f a"
    subprocess.call([f]) #opens alphamelts

    try:
        for files in os.walk("Whole_Planet_MELTS_Files"):
            pyautogui.typewrite('1'), pyautogui.press('enter') #enter MELTS file
            pyautogui.typewrite(str(current_file)) #this is not correct, but need to interate through every file and enter the name into the section
            pyautogui.typewrite("5"), pyautogui.press('enter') #set fO2
            pyautogui.typewrite('4'), pyautogui.press('enter') #IW
            pyautogui.typewrite('-1.4'), pyautogui.press('enter') #IW-1.4
            pyautogui.typewrite('8'), pyautogui.press('enter') #select phase to omit
            pyautogui.typewrite("alloy-liquid"), pyautogui.press('enter') #selets liquid alloy phase
            pyautogui.typewrite("0"), pyautogui.press('enter') #omits liquid alloy phase from calculations
            pyautogui.typewrite("4"), pyautogui.press('enter') #begins calculations
            time.sleep(30) #wait for calculations to finish
            pyautogui.typewrite("0") #shuts down alphamelts and generates output file

            subprocess.check_call(f)
            if subprocess.check_call(f) == 0:
                print "Calculation Successful.  Moving on..."
            elif subprocess.check_call(f) == 1:
                print "Calculation Unsucessful.  Moving on..."
                #then, write star name with success/failure status to another spreadsheet and continue on tot he next file in the directory
            else:
                print "alphaMELTS did not run correctly.  Moving on..."


            #this copies files to a new folder, and then it renames the file
            def copy_rename(): #alphaMELTS_tbl.txt is the output file, want to copy, rename, and save that file in the completed outputs directory
                src_dir = os.curdir
                dst_dir = os.path.join("Completed_Whole_Planet_MELTS_Files")
                src_file = os.path.join(src_file, 'alphaMELTS_tbl.txt')
                shutil.copy(src_file, dst_dir)

                dst_file = os.path.join(dist_dir, 'alphaMELTS_tbl.txt')
                os.rename(dst_file, str(current_file_minus_txt) + '_COMPLETED_BSP_.txt.') #renames alphaMELTS_tbl.txt to the name of the MELTS file from which it was generated
                #http://techs.studyhorror.com/python-copy-rename-files-i-122

            copy_rename()    #executes the copy/rename function described above




    except:
        print "Error.  BSP Calculation Unsuccessful."
        initialization()



#automates alphaMELTS for basalt calculations
def run_basalt():
    if not os.path.exists('Completed_BSP_MELTS_Files'):
        os.mkdir('Completed_BSP_MELTS_Files') #create a directory for successful output files
    os.path.join('Completed_Whole_Planet_MELTS_Files')

    current_file = os.path.basename("Whole_Planet_MELTS_Files")  #for use when copying and renaming the alphaMELTS_tbl.txt file
    current_file_minus_txt = current_file[:-3]

    f = "run_alphamelts.command -f a"
    subprocess.call([f]) #opens alphamelts

    try:
        for files in os.walk('BSP_MELTS_Files'):
            current_file = os.path.basename("BSP_Melts_Files")
            pyautogui.typewrite('1'), pyautogui.press('enter') #enter MELTS file
            pyautogui.typewrite(str(current_file)) #this is not correct, but need to interate through every file and enter the name into the section
            pyautogui.typewrite('5'), pyautogui.press('enter') #select fO2
            pyautogui.typewrite('3'), pyautogui.press('enter') #QFM
            pyautogui.press('+1.5'), pyautogui.press('enter') #QFM+1.5
            pyautogui.typewrite('8'), pyautogui.press('enter') #surpress phase
            pyautogui.typewrite("alloy-liquid"), pyautogui.press('enter') #phase alloy-liquid
            pyautogui.typewrite("0"), pyautogui.press('enter') #surpress phase
            pyautogui.typewrite("3"), pyautogui.press('enter') #run batch calculation
            pyautogui.typewrite('1'), pyautogui.press('enter') #superliquidus
            pyautogui.typewrite('1'), pyautogui.press('enter') #mass percent type calculation
            pyautogui.typewrite('0.08'), pyautogui.press('enter') #8 percent by mass
            time.sleep(20) #get code to wait for calculation to finish before next command.  Better way to do this, like check_output???
            pyautogui.typewrite('4'), pyautogui.press('enter') #begin calculations
            time.sleep(30) #wait for calculations to finish
            pyautogui.typewrite("0") #shuts down alphamelts and generates output file

            subprocess.check_call(f)
            if subprocess.check_call(f) == 0:
                print "Calculation Successful.  Moving on..."
            elif subprocess.check_call(f) == 1:
                print "Calculation Unsucessful.  Moving on..."
            else:
                print "alphaMELTS did not run correctly.  Moving on..."

            def copy_rename(): #alphaMELTS_tbl.txt is the output file, want to copy, rename, and save that file in the completed outputs directory
                src_dir = os.curdir
                dst_dir = os.path.join("Completed_Whole_Planet_MELTS_Files")
                src_file = os.path.join(src_file, 'alphaMELTS_tbl.txt')
                shutil.copy(src_file, dst_dir)

                dst_file = os.path.join(dist_dir, 'alphaMELTS_tbl.txt')
                os.rename(dst_file, str(current_file_minus_txt) + '_COMPLETED_BASALT_.txt.') #renames alphaMELTS_tbl.txt to the name of the MELTS file from which it was generated
                #http://techs.studyhorror.com/python-copy-rename-files-i-122

            copy_rename()   #executes the copy/rename function described above

    except:
        print "Error.  Basalt Calculation Unsuccessful."
        #write failure to file and continue on to next file in the directory


def core_removal(): #take chemistry from rewritten alphaMELTS_tbl.txt files and subtracts out total moles of Fe/Ni alloy, then renormalizes chemistry to 100 percent





#the code actually begins running here
    initialization()