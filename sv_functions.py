####################################################
#                  Revision: 1.1                   #
#              Updated on: 11/13/2015              #
#                                                  #
# What's new:                                      #
#                                                  #
# CONDITIONAL FORMATING:                           #
#                                                  #
#   Font color will change if it matches given     # 
#   conditions.                                    #
#                                                  #
#   Added two functions to support it:             #
#                                                  #
#       1.) add_conditional_formatting_sv          #
#       2.) perform_CF_on_list                     #
#                                                  #
####################################################

####################################################
#                  Revision: 1.0                   #
#              Updated on: 11/06/2015              #
####################################################

####################################################
#                                                  #
#   This script contains all functions usefull     #
#   to generate a report for SV Test               #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################

import pandas
import numpy as np

class SV_Functions:

    ###################################
    # This function will take data and 
    # extract of that particular column 
    # given by "column_index" and 
    # particular indices of that columns
    # given by "index_list"
    ###################################
    def extract_given_indices_data(data, column_index, index_list):
    
        temp = np.array(data[column_index][:])
        extracted_list = [ float((temp[index_list[j]].tolist())[0]) 
                    for j in range(len(index_list))]

        return extracted_list

    
    
    #####################################
    # Filter of files only with  
    # required string(s) in file names
    #
    # If "selection" == 0 include string
    #
    # If "selection" == 1 exclude string 
    #####################################
    def filter(file_list, str_name, selection):
        
        if selection == 0: # INCLUDE
            file_list =[file_list[i] for i in range(len(file_list)) if 
                       str_name in file_list[i]]
        
        if selection == 1: # EXCLUDE
            file_list =[file_list[i] for i in range(len(file_list)) if 
                       str_name not in file_list[i]]
        return file_list
        
        
        
    #################################
    # Sorting files so that we 
    # extract data directly. And each 
    # drives will be a sequence. 
    #################################
    def sort_files(file_list):
        sorted_files = sorted(file_list, key=lambda x: int(x.split('_')[0]))

        return sorted_files
    
     

    #############################################
    # This function will search specific
    # string "search_str" in a single column
    # and return their indices
    #############################################
    def find_string_indices(search_str, single_column):
        str_indices = [i for i in range(len(single_column)) 
                          if search_str in str(single_column[i]) ]
        return str_indices
        
        
            
    #######################
    # Finding indices of 
    # 1st and 2nd Test
    #######################
    def find_first_second_test_indices(phydrive_indices):

        # Finding Drives used for each tests
        disk_no = [ i for i in range(len(phydrive_indices)) 
                   if phydrive_indices[i] - phydrive_indices[i-1] > 1  ]

        #print(disk_no)

        # Indices of 1st test
        test_1_indices = phydrive_indices[0:disk_no[0]]
        #print(test_1_indices)

        # Indices of 2nd test
        test_2_indices = phydrive_indices[disk_no[0]:(2*disk_no[0])+1]
        #print(len(test_2_indices), test_2_indices)
        
        return [disk_no, test_1_indices, test_2_indices]

    
    
    #############################################
    # This function will search Time Stamp info 
    # from the given file "file_name". 
    #
    # First time stamp should be present in 
    # 1st column and 6th row of given .csv
    # file.
    #############################################
    def extract_time_stamps(dir_path, file_name):
        import os
        import sys
        
        ###################################
        #  Importing from other Directory
        ###################################
        os.chdir('..')
        c_path = os.getcwd()
        sys.path.insert(0, r''+str(c_path)+'/Common Scripts')
        
        import report_functions
        rf= report_functions.Report_Functions

        sys.path.insert(0, r''+str(c_path)+'/SV')
        
        #################################
        # Extractting 1st Time Stamp 
        # info
        #################################
        csv_data1= pandas.read_csv( r''+str(dir_path)+'\\'
                                   + str(file_name)
                                   , nrows=6
                                   , header= None )

        time_stamp_1st= csv_data1[0][5]


        #################################
        # Extractting 2nd Time Stamp 
        # info
        #################################
        csv_data2= pandas.read_csv( r''+str(dir_path)+'\\'
                                   + str(file_name)
                                   , skiprows=13
                                   , header= None )

        [ts_count, ts_indices]= rf.find_string(csv_data2
                                                , 0, 0
                                                ,'\'Time Stamp')

        time_stamp_2nd= csv_data2[0][ts_indices[0]+1]

        final_time_stamps=np.hstack((time_stamp_1st
                                    ,time_stamp_2nd))
                                    
        final_time_stamps=final_time_stamps.tolist()

        final_time_stamps=['Time Stamp']+final_time_stamps
        
        return final_time_stamps
        
        
        
    #################################
    # Extractting 1st Time Stamp 
    # info
    #################################    
    def extract_1st_time_stamp(dir_path, file_name):
        import os
        import sys
        
        ###################################
        #  Importing from other Directory
        ###################################
        os.chdir('..')
        c_path = os.getcwd()
        sys.path.insert(0, r''+str(c_path)+'/Common Scripts')
        
        import report_functions
        rf= report_functions.Report_Functions

        sys.path.insert(0, r''+str(c_path)+'/SV')
        
        #################################
        # Extractting 1st Time Stamp 
        # info
        #################################
        csv_data1= pandas.read_csv( r''+str(dir_path)+'\\'
                                   + str(file_name)
                                   , nrows=6
                                   , header= None )

        return [csv_data1[0][5]]

    
    
    
    ##########################
    # To extract drive numbers
    # from .csv file
    #########################
    def extract_drive_no(col_1, test_1_indices):
        
        import re

        drive_nos = [] 
        for j in range(len(test_1_indices)):

            temp_no = re.findall(r'\d+', (col_1[test_1_indices[j]]))
            temp_no = 'Drive#' + str(temp_no[0])
            drive_nos.append(temp_no)
        drive_nos      

        final_drives = ['System'] + drive_nos
        return final_drives


    
    #####################
    # Extract Test names 
    # for different test
    # Say: 4k & 512k
    # from the 1st file.
    # 
    # It is assumed that
    # Test Description
    # remains same
    # throughout all
    # baseline and
    # performance files.
    #####################
    def extract_test_description(file_data, file_test_name):
    
        ###################################
        #  Importing from other Directory
        ###################################
        import os
        import sys
        
        os.chdir('..')
        c_path = os.getcwd()
        sys.path.insert(0, r''+str(c_path)+'/Common Scripts')
        
        import report_functions
        rf= report_functions.Report_Functions

        sys.path.insert(0, r''+str(c_path)+'/SV')
    
        j=0
        [worker_count, worker_indices]= rf.find_string(file_data,0,0,'WORKER')
        
        test_name_test_type = [(file_data[2][worker_indices[j]])+', '+str(file_test_name) 
                      for j in range(len(worker_indices)) ] 
        
        just_test_name= [(file_data[2][worker_indices[j]]) 
                      for j in range(len(worker_indices)) ] 
                      
        test_name_test_type = ['Test Description'] + test_name_test_type
    
        return [test_name_test_type, just_test_name]
            

            
    ##########################
    # extract name from 
    # file name.
    #
    # If "selector" == 0:
    #   Extract Baseline names
    #
    # If "selector" == 1:
    #   Extract Performance names
    ##########################
    def extract_file_name(temp, selector):
        if int(selector) == 0:
            for i in range(len(temp)):
                if temp[i] == '_':
                    temp_index = i
                    break
            temp_name = temp[temp_index+1:]
            bs_index = temp_name.find('Baseline',0, len(temp_name))

        elif int(selector) == 1:
            for i in range(len(temp)):
                if temp[i] == '_':
                    temp_index = i
                    break
            temp_name = temp[temp_index+1:]
            bs_index = temp_name.find('Perf',0, len(temp_name))

        return temp_name[:bs_index-1]        
    
            
            
    ####################################
    # This function will extract all 
    # required data. That is 
    # 
    #   1.) IOps of 1st & 2nd test
    #   2.) Errors of 1st & 2nd test
    #   3.) Drive Numbers 
    #   4.) Time Stamps
    #   5.) Test Description
    ####################################
    
    def extract_baseline_data(dir_path, file_name, selector):
    
        import pandas
        import numpy as np
        import os
        import sys
        
        #print(file_name)
        file_test_name = SV_Functions.extract_file_name(file_name, selector)
        
        bs_data = pandas.read_csv(r''+str(dir_path)+'\\'
                                   + str(file_name)
                                   , skiprows=13
                                   , header= None )

        # using 1st column to find useful string i.e "Worker" 
        # & "Physical drive" numbers
        
        #bs_data = np.array(bs_data)gar
        col_1 = bs_data[1][:] # 1st column
        
        
        # worker indices
        worker_indices = SV_Functions.find_string_indices('Worker', col_1)
        
        
        # PHYSICALDRIVE: indices
        phydrive_indices = SV_Functions.find_string_indices('PHYSICALDRIVE:', col_1)
        #print(phydrive_indices)
        
        
        #######################
        # Finding indices of 
        # 1st and 2nd Test
        #######################
        [disk_no, test_1_indices, test_2_indices] = SV_Functions.find_first_second_test_indices(phydrive_indices)
        #print(disk_no)

        ###################################
        #  Importing from other Directory
        ###################################
        #os.chdir('..')
        c_path = os.getcwd()
        
        ##################################
        # Inserting "Common Scripts" path 
        # to import and use some of 
        # its functions
        ##################################
        sys.path.insert(0, r''+str(c_path)+'/Common Scripts')

        import report_functions
        rf = report_functions.Report_Functions
        
        ##################################
        # Jumping back to "SV" directory
        # This is our original directory
        ##################################
        sys.path.insert(0, r''+str(c_path)+'/SV')
        #print(os.getcwd())
        
        # Finding IOps's Column number
        [iops_counts, iops_index]= rf.find_string(bs_data
                                                  , 0
                                                  , 1
                                                  , 'IOps')
                                                   
        #print(iops_index)

        # Finding Error's Column number
        [error_counts, error_index]= rf.find_string(bs_data
                                                  , 0
                                                  , 1
                                                  , 'Errors')
                                                  
        
        ####################
        # IOps
        ###################
        # IOps 1st test
        ###################
        iops_1st = SV_Functions.extract_given_indices_data(bs_data
                                                          , iops_index
                                                          , test_1_indices)
        
        ###################
        # IOps 2nd test
        ###################
        iops_2nd = SV_Functions.extract_given_indices_data(bs_data
                                                        , iops_index
                                                        , test_2_indices)
        
        ########################
        # For System/Worker IOps
        ########################
        iops_system = SV_Functions.extract_given_indices_data(bs_data
                                                            , iops_index
                                                            , worker_indices)
        #######################
        # Appending System IOps
        # to IOps list
        #######################
        #from statistics import mean
        #final_iops_1st = [round(mean(iops_1st),6)] + iops_1st

        #final_iops_2nd = [round(mean(iops_2nd),6)] + iops_2nd
        
        final_iops_1st = [iops_system[0]] + iops_1st

        final_iops_2nd = [iops_system[1]] + iops_2nd
        
        ####################
        # Errors
        ####################
        ###################
        # Errors 1st test
        ###################

        errors_1st = SV_Functions.extract_given_indices_data(bs_data
                                                            , error_index
                                                            , test_1_indices)

        ###################
        # Errors 2nd test
        ###################

        errors_2nd = SV_Functions.extract_given_indices_data(bs_data
                                                            , error_index
                                                            , test_2_indices)
                                                
        ###########################                                        
        # For System/Worker Errors
        ###########################                                        
        errors_system = SV_Functions.extract_given_indices_data(bs_data
                                                                , error_index
                                                                , worker_indices)

        #########################
        # Appending System Errors
        # to Errors list
        #########################
        final_errors_1st = [errors_system[0]] + errors_1st
        #final_errors_1st

        final_errors_2nd = [errors_system[1]] + errors_2nd
        
        
        ######################
        # Get Drive numbers
        ######################
        final_drives = SV_Functions.extract_drive_no(col_1
                                                    , test_1_indices)
                                                    
        
 
        ######################
        # Get test description
        ######################
        
        [test_name_list, just_test_name] = SV_Functions.extract_test_description(bs_data, file_test_name)
        
        final_time_stamps = SV_Functions.extract_time_stamps(dir_path, file_name)
        
        
        return [final_iops_1st, final_errors_1st
                , final_iops_2nd, final_errors_2nd
                , final_drives, test_name_list
                , final_time_stamps, disk_no, just_test_name]
                
                
                
    ##################################
    # This Function will return an
    # Alphabet related to its Numbers
    #
    # Example rank(1) = A
    #         rank(26) = Z
    #
    # Note that number should be less
    # or equal to 26. Or else it will
    # throw a KeyError
    ##################################
    def rank(x):
        import string
        d = dict((n%26+1,letr) for n,letr in enumerate(string.ascii_letters[0:52]))
        return d[x]     
        
    
    
    ######################################
    # PERFORMANCE FILES
    ###########################
    # extract Office vibe data
    ###########################
    def extract_office_vibe_single_file(dir_path, file_name, selector):
        
        import pandas
        import numpy as np
        import os
        import sys

        file_test_name = SV_Functions.extract_file_name(file_name, selector)

        bs_data = pandas.read_csv(r''+str(dir_path)+'\\'
                                   + str(file_name)
                                   , skiprows=13
                                   , header= None )

        # using 1st column to find useful string i.e "Worker" 
        # & "Physical drive" numbers

 
        col_1 = bs_data[1][:] # 1st column


        # worker indices
        worker_indices = SV_Functions.find_string_indices('Worker', col_1)


        # PHYSICALDRIVE: indices
        test_1_indices = SV_Functions.find_string_indices('PHYSICALDRIVE:', col_1)
 

        ###################################
        #  Importing from other Directory
        ###################################

        c_path = os.getcwd()
        
        ##################################
        # Inserting "Common Scripts" path 
        # to import and use some of 
        # its functions
        ################################## 
        sys.path.insert(0, r''+str(c_path)+'/Common Scripts')
   
        import report_functions
        rf = report_functions.Report_Functions

        ##################################
        # Jumping back to "SV" directory
        # This is our original directory
        ##################################
        sys.path.insert(0, r''+str(c_path)+'/SV')


        # Finding IOps's Column number
        [iops_counts, iops_index]= rf.find_string(bs_data
                                                  , 0
                                                  , 1
                                                  , 'IOps')



        # Finding Error's Column number
        [error_counts, error_index]= rf.find_string(bs_data
                                                  , 0
                                                  , 1
                                                  , 'Errors')


        ####################
        # IOps
        ###################
        # IOps 1st test
        ###################
        iops_1st = SV_Functions.extract_given_indices_data(bs_data
                                                          , iops_index
                                                          , test_1_indices)



        ########################
        # For System/Worker IOps
        ########################
        iops_system = SV_Functions.extract_given_indices_data(bs_data
                                                            , iops_index
                                                            , worker_indices)
        #######################
        # Appending System IOps
        # to IOps list
        #######################

        final_iops_1st = [iops_system[0]] + iops_1st



        ####################
        # Errors
        ####################
        ###################
        # Errors 1st test
        ###################

        errors_1st = SV_Functions.extract_given_indices_data(bs_data
                                                            , error_index
                                                            , test_1_indices)

                                               

        ###########################                                        
        # For System/Worker Errors
        ###########################                                        
        errors_system = SV_Functions.extract_given_indices_data(bs_data
                                                                , error_index
                                                                , worker_indices)

        #########################
        # Appending System Errors
        # to Errors list
        #########################
        final_errors_1st = [errors_system[0]] + errors_1st



        ######################
        # Get Drive numbers
        ######################
        final_drives = SV_Functions.extract_drive_no(col_1
                                                    , test_1_indices)



        ######################
        # Get test description
        ######################

        [test_name_list, just_test_name] = SV_Functions.extract_test_description(bs_data, file_test_name)

        final_time_stamps = SV_Functions.extract_1st_time_stamp(dir_path, file_name)
        

        return [final_iops_1st, final_errors_1st
                , final_drives, test_name_list
                , final_time_stamps, just_test_name]
                
        

    def calculate_n_write_stats(worksheet, row_new, row, col, format):    
        # Calculating Maximum from Degradation list
        worksheet.write_formula(row_new, col, '=MAX($'+
                                       str(SV_Functions.rank(col+1))+str(row+2)+':$'+
                                       str(SV_Functions.rank(col+1))+str(row_new)+ ')', format)
                                    
        # Calculating Minimum from Degradation list
        worksheet.write_formula(row_new+1, col, '=MIN($'+
                                       str(SV_Functions.rank(col+1))+str(row+2)+':$'+
                                       str(SV_Functions.rank(col+1))+str(row_new)+ ')', format)

        # Calculating Average from Degradation list                            
        worksheet.write_formula(row_new+2, col, '=AVERAGE($'+
                                       str(SV_Functions.rank(col+1))+str(row+2)+':$'+
                                       str(SV_Functions.rank(col+1))+str(row_new)+ ')', format)

        # Calculating Standard Deviation from Degradation list
        worksheet.write_formula(row_new+3, col, '=STDEV($'+
                                       str(SV_Functions.rank(col+1))+str(row+2)+':$'+
                                       str(SV_Functions.rank(col+1))+str(row_new)+ ')', format)
                                       
        return 0

    
    ############################################
    # CONDITIONAL FORMATING
    #
    # This function will change 
    # the font color conditionally.
    #
    #   RED:
    #    if "red_value"
    #    is greater than and equal to value(>=) 
    #    Cell value.
    #
    #   BLUE:
    #    if "blue_value"
    #    is less than and equal to value(<=) 
    #    Cell value.
    #
    #   "selection" used to select row-wise
    #    or column-wise formatting
    #   
    #    if 
    #       selection = 0: Row Wise formatting
    #       selection = 1: Column Wise formatting
    #                      with same Font color 
    #                      scheme
    ############################################
    def add_conditional_formatting_sv(ssef, selection, workbook, worksheet, row, col, list_length, blue_value, red_value):
        
        #ssef = SSE_Functions
        
        #list_length = len(list)
        
        # For RED FONT
        red_format = workbook.add_format({
                                            'font_color': 'red'
        
                                        })
        
        # For BLUE FONT
        blue_format = workbook.add_format({
                                            'font_color': 'blue'
                                         })
        if selection == 0: # ROW_WISE                                 
            worksheet.conditional_format('$'+str(ssef.rank(col+1))+ '$' + str(row+1) + ':$' +str(ssef.rank(col+1))+ '$' +str(row+list_length),

                                            {
                                                'type': 'cell',
                                                'criteria': '>=',
                                                'value': red_value,
                                                'format': red_format,
                                            }
                                        )
            
            worksheet.conditional_format('$'+str(ssef.rank(col+1))+ '$' + str(row+1) + ':$' +str(ssef.rank(col+1))+ '$' +str(row+list_length),
                                            {
                                                'type': 'cell',
                                                'criteria': '<=',
                                                'value': blue_value,
                                                'format': blue_format,
                                            }
                                        )
        elif selection == 1: # COLUMN_WISE                                 
            worksheet.conditional_format('$'+str(ssef.rank(col+1))+ '$' + str(row+1) + ':$' +str(ssef.rank(col+list_length))+ '$' +str(row+1),

                                            {
                                                'type': 'cell',
                                                'criteria': '>=',
                                                'value': red_value,
                                                'format': red_format,
                                            }
                                        )
            
            worksheet.conditional_format('$'+str(ssef.rank(col+1))+ '$' + str(row+1) + ':$' +str(ssef.rank(col+list_length))+ '$' +str(row+1),
                                            {
                                                'type': 'cell',
                                                'criteria': '<=',
                                                'value': blue_value,
                                                'format': blue_format,
                                            }
                                        )
        return [workbook, worksheet]
    
        
    ############################
    # Conditional Formatting
    # if 4K, use -5, 5 for 
    # formatting
    #
    # if 512K, use -10, 10 for
    # formatting
    ############################
    def perform_CF_on_list(ssef,    workbook, worksheet, row, col, list, test_name):
        if '4' in str(test_name):
            # System
            [workbook, worksheet] = SV_Functions.add_conditional_formatting_sv(ssef, 0, workbook, worksheet, row, col, len([list[0]]), -10, 10)                      
            
            # Drives
            [workbook, worksheet] = SV_Functions.add_conditional_formatting_sv(ssef, 0, workbook, worksheet, row+1, col, len(list[1:])+4, -5, 5)     
            
        elif '512' in str(test_name):                      
            # System
            [workbook, worksheet] = SV_Functions.add_conditional_formatting_sv(ssef, 0, workbook, worksheet, row, col, len([list[0]]), -15, 15)                      
            
            # Drives
            [workbook, worksheet] = SV_Functions.add_conditional_formatting_sv(ssef, 0, workbook, worksheet, row+1, col, len(list[1:])+4, -10, 10)                      
            
        return [workbook, worksheet]
    

#############################
#           END             #
#############################        