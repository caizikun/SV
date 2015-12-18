####################################################
#                  Revision: 1.1                   #
#              Updated on: 11/13/2015              #
#                                                  #
# What's new:                                      #
#       Calling some functions from                #
#       "sv_functions.py" for conditional          #
#       formatting.
#                                                  #
# CONDITIONAL FORMATING:                           #
#                                                  #
#   Font color will change if it matches given     # 
#   conditions.                                    #
#                                                  #
#   Added two functions in sv_functions.py         #
#   to support it:                                 #
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
#   This script contains a functions usefull       #
#   to extract data from from the "sorted_files"   #
#   and write in a Excel file.                     # 
#                                                  #
#   Excel File name is provided by User            #
#   as "report_name".                              #
#                                                  #
#   "chassis_name" is used to title the chart      #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################


class Extract_Create_Baseline:

    #########################################
    # Function will extract data from 
    # given Baseline files from main 
    # file. Manipulate it and write it
    # in worksheet "ws_1st_name".
    # 
    # Using this data, it will plot and
    # add chart to "ws_2nd_name" worksheet
    #########################################
    def generate_excel_baseline(c_path, sorted_files, file_dir_perf, workbook
                       , chassis_name, ws_1st_name, ws_2nd_name, selector, title):
        
        import os
        import xlsxwriter
        
        import sv_functions
        svf = sv_functions.SV_Functions

        
        ###################################
        #  Importing from other Directory
        ###################################
        import sys
        import os
        import numpy as np

        ##################################
        # Inserting "SSE" path to import 
        # and use some of its functions
        ##################################        
        os.chdir('..')
        c_path = os.getcwd()
        sys.path.insert(0, r''+str(c_path)+'/SSE')

        import sse_functions
        ssef= sse_functions.SSE_Functions

        ##################################
        # Jumping back to "SV" directory
        # This is our original directory
        ##################################
        sys.path.insert(0, r''+str(c_path)+'/SV')                               
                
                
        # length/number of sorted baseline files
        sf_len = len(sorted_files) 
        
        
        
        

                                       
        #############################
        # Create Baseline Worksheet
        #############################
        worksheet = workbook.add_worksheet(''+str(ws_1st_name))
        
        # Set column width to 30 points
        worksheet.set_column('A:Z', 30)

        
        ############################
        # Create & Formate 
        # "Baseline" worksheet in 
        # a Report.
        ############################
        
        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
                                            'border': 1,
                                            'align': 'left',
                                            'valign': 'left',
                                            'fg_color': 'yellow',
                                            'bold': True,
                                            'font_name':'Arial',
                                            'font_size':14
                                           })

        
        # Add a bold format to use to highlight cells.
        bold_14 = workbook.add_format({
                                        'bold': True, 
                                        'font_name':'Arial',
                                        'font_size':14
                                      })


      
        bold_12 = workbook.add_format({ 
                                        'bold': True,
                                        'font_name':'Arial',
                                        'font_size':12
                                      })

                                      
        bold_10 = workbook.add_format({
                                        'bold': True, 
                                        'font_name':'Arial',
                                        'font_size':10,
                                        'align': 'center'
                                      })
        
        bold_10_l = workbook.add_format({
                                        'bold': True, 
                                        'font_name':'Arial',
                                        'font_size':10,
                                        'align': 'left'
                                      })
       
            
        regular_10 = workbook.add_format({
                                          'font_name':'Arial',
                                          'font_size':10,
                                          'align': 'center'
                                        })

        # Font: Arial, Font_size:10 & left Aligned
        regular_10_l = workbook.add_format({
                                            'font_name':'Arial',
                                            'font_size':10,
                                            'align': 'left'
                                           })
        
        # Set border for certain fonts
        regular_10.set_border(1)
        bold_10.set_border(1)
        bold_10_l.set_border(1)
        regular_10_l.set_border(1)
        
        format1 = workbook.add_format({'bold':True, 'align': 'center'})
        format1.set_num_format('0.0')
        format1.set_border()
        
        
        #########################################
        # Merge 1st & 2nd rows, over columns.
        #########################################
        worksheet.merge_range('A1:Z1', 'BASELINE TEST RESULTS', merge_format)
        worksheet.merge_range('A2:Z2', ''+str(title), merge_format)
        
        
        #########################################
        # Including Title for IOps Result,
        # "col" & "row" indicates position 
        # in the Excel sheet.
        #########################################
        col = 0
        row = 3
        worksheet.write(row, col, 'IOps Results', bold_12)
        
        
        ################################
        # Writing "All Drives Baseline"                          
        ################################                                              
        worksheet1 = workbook.add_worksheet(''+str(ws_2nd_name))                          
        

        #########################################
        # Chart Settings
        #########################################        
        
        # chart type
        chart = workbook.add_chart({'type': 'line'})
        
        # chart scale
        chart.set_size({'x_scale': 3, 'y_scale': 2.5}) 
        
        # chart title
        chart.set_title({
                         'name':''+str(chassis_name)
                         +' Basline Testing'\
                         '\nwith ' +str(title)+
                         '\nIOMeter Test Results by Drive '
                        })
        
        
        # chart y axis range and Axis name                        
        chart.set_y_axis({
                           'name': 'IOps % Degradation',
                           'min': -15,
                           'max': 15
                         })
                         
                         
        # chart x axis range, label position, gridlines, and Axis name                        
        chart.set_x_axis({
                           'label_position': 'low',
                           'position_axis': 'on_tick',
                           'min': 1,
                           'major_gridlines': {
                                                'visible': True,
                                                'line': {'width': 0.5}
                                              }
                         })

        

                         
        ###################################
        # This loop will read and extract  
        # Baseline file one by one. 
        #
        # It will write IOps data, 
        # Degradation, IOMeter Errors for 
        # both tests.
        ###################################
        for sf in range(sf_len):
            
            ######################################
            # This function will extract all data 
            # from Single .csv Baselie file.
            ######################################            
            [final_iops_1st, final_errors_1st
            , final_iops_2nd, final_errors_2nd
            , final_drives, test_name_list
            , final_time_stamps, disk_no
            , just_test_name] = svf.extract_baseline_data(file_dir_perf,
                                                          sorted_files[sf],
                                                          selector)
            #print(len(final_iops_1st))
            
            
            ######################################                                                          
            # Joining/Modifying List so that we 
            # can directly write it in 1st Column
            ######################################
            
            # 1st column,IOps
            col_1st = [final_time_stamps[0]] + [test_name_list[0]] + final_drives
            
            # 2nd column,IOps
            col_2nd = [final_time_stamps[1]] + [test_name_list[1]] 
            
            # 3rd column,IOps
            col_3rd = [final_time_stamps[2]] + [test_name_list[2]]
            
            
            
            ##############################
            # We need to write 1st column
            # just once. So if it is a 
            # 1st file. Write 1st, 2nd, 
            # and 3rd column.
            ##############################
            if sf == 0:
            
                # store data of 1st file to find degradation value later
                file1_file_1st = final_iops_1st 
                file1_file_2nd = final_iops_2nd

                
                ######################################################
                # creating an empty list of zeros, which will finally
                # hold row-wise summation of All .csv files.
                # Which is used to calculate the Average
                ######################################################
                
                # 1st Test
                null_array_1st = np.zeros(int(disk_no[0])+1)
                null_array_1st = null_array_1st.tolist() 
                #print(len(null_array_1st))
                # 2nd Test
                null_array_2nd = null_array_1st
                

                
                #########################################
                # Zeros array to calculate average
                #########################################

                # Summing up 1st test data
                null_array_1st = np.sum([final_iops_1st, null_array_1st], axis = 0)

                # Summing up 2nd test data
                null_array_2nd = np.sum([final_iops_2nd, null_array_2nd], axis = 0)
                      
                
                
                #############################################                    
                # Writing 1st column, IOps, Left justified     
                # Originally row = 3, for IOps title.
                # Now row =4, to write it from next 
                # row, col = 0 (1st column)
                #############################################
                
                row+= 1 
                ssef.write_excel_data( worksheet
                                      , col_1st, row
                                      , col, 0, regular_10_l)
                
                
                # Writing 2nd column, center justified
                col+=1  # Moving to next Column
                row1 = row + 2 
                
                #############################################                 
                # this function will write given list 
                # of strings row-wise
                #############################################                
                ssef.write_excel_data( worksheet, col_2nd, row, 
                                       col, 0, regular_10)
                
                
                #######################################
                # this function will write given list 
                # of numbers, row-wise                      
                #######################################
                ssef.write_excel_data_float( worksheet
                                           , final_iops_1st, row1
                                           , col, 0, regular_10)
                             
                             
                # Writing 3rd column
                col+=1                      
                ssef.write_excel_data( worksheet, col_3rd, row,
                                       col, 0, regular_10)
                                      
                ssef.write_excel_data_float( worksheet, final_iops_2nd, 
                                             row1, col, 0, regular_10)
            
            
            
            
                ####################################
                # Writing degradation data 
                ####################################
                # calculate next row to write data
                #####################################
                
                ###################################################
                #
                # starting row (row2)=   row1(staring IOps row) 
                #                       + (number of disks + system) 
                #                       + 4 Blank Rows  
                #
                ####################################################
                row2 = row1 + (disk_no[0]+1) + 4 
                col2 = 0 # 1st column 
                
                # Degradation title
                worksheet.write(row2, col2, 'IOps, % degradation, compared to initial baseline test', bold_12)
                
                
                # Writing 1st column, Degradation     
                row2 += 1
                ssef.write_excel_data( worksheet, 
                                       col_1st[2:],
                                       row2, 
                                       col2, 
                                       0, 
                                       regular_10_l)
            
            
                #####################################
                # As degradation is calculated with 
                # respect to 1st file's values
                # leave 2nd and 3rd column blank.
                #
                # Write Statistical titles in 3rd column
                #########################################
                col2 += 2
                row2_new = row2 + (disk_no[0]+1)
                
                # To write Statistical data titles
                stat_list = ['High', 'Low', 'Average', '1 σ']
                
                # Writing stat list
                ssef.write_excel_data( worksheet
                                      , stat_list, row2_new
                                      , col2, 0, bold_10_l)
                
                
                
                ####################################
                # Writing IOMeter Error data 
                ####################################
                # calculate next row to write data
                #####################################
                
                ###################################################
                #
                # starting row (row3)=   row2(staring Degradation row) 
                #                       + (number of disks + system) 
                #                       + 6 Blank Rows  
                #
                ####################################################
                row3 = row2 + (disk_no[0]+1) + 6  
                
                # Error Title
                col3 = 0 
                worksheet.write(row3, col3, 'IOMeter Errors', bold_12)
                
                # Writing 1st column, Drives #     
                row3 += 1
                ssef.write_excel_data( worksheet
                                      , col_1st[2:], row3
                                      , col3, 0, regular_10_l)
                
                # Writing 2nd column, 1st Test Errors
                col3+=1  
                      
                ssef.write_excel_data( worksheet
                                      , final_errors_1st, row3
                                      , col3, 0, regular_10)
                
                # Writing 3rd column, 2nd Test Errors                
                col3+=1                      
                ssef.write_excel_data_float( worksheet
                                          , final_errors_2nd, row3
                                          , col3, 0, regular_10)

            
            
            
            
            
            ##############################
            # We need to write 1st column
            # just once. So if it is a 
            # 1st file. Write 1st, 2nd, 
            # and 3rd column.
            #
            # Else just write 2nd and 3rd 
            # column from 2nd and rest of 
            # files.
            ##############################
                                         
            else:
            
                ##################################################
                # Null array to calculate average
                #
                # Summing up 1st test data, later it will used to 
                # calculate the average
                ##################################################
                null_array_1st = np.sum([final_iops_1st, null_array_1st], axis = 0)

                # Summing up 2nd test data
                null_array_2nd = np.sum([final_iops_2nd, null_array_2nd], axis = 0)
         
         
            
                ######################################
                # this function will write given list 
                # of strings row-wise
                ######################################
                row1 = row + 2 
                
                
                ##################################
                # 2n + 1 series, It will generate
                # odd numbers.
                #
                # In our case, as sf starts from 1,
                # we will get:
                #              3,5,7,...(2n+1)
                #
                # Which is 4th, 6th,...(2n+2) columns
                # from Excel point of view. As it 
                # starts numbering from 1.
                #
                # So in Excel worksheet, even columns 
                # will have 1st test data.
                #####################################
                col = 1 + (sf*2)   
                
                
                ########################################
                # this function will write given list 
                # of numbers, row-wise 
                #
                #           Writing 2nd column
                ########################################
                
                # just titles: Time stamp & Test Description 
                ssef.write_excel_data( worksheet
                                      , col_2nd, row
                                      , col, 0, regular_10)
                                      
                # writing data exracted from Baseline file(s)
                ssef.write_excel_data_float( worksheet
                                            , final_iops_1st, row1
                                            , col, 0, regular_10)
                

                
                ##################################
                # 2n + 2 series, It will generate
                # even numbers.
                #
                # In our case, as sf starts from 1,
                # we will get:
                #              4,6,8,...(2n+2)
                #
                # Which is 5th, 7th,...(2n+3) columns
                # from Excel point of view. As it 
                # starts numbering from 1.
                #
                # So in Excel worksheet, odd columns 
                # will have 2nd test data from Same 
                # file.
                #####################################
                col = 2 + (sf*2)
                
                # Writing 3rd column
                # just titles: Time stamp & Test Description 
                ssef.write_excel_data( worksheet
                                      , col_3rd, row
                                      , col, 0, regular_10)
                
                # writing data exracted from Baseline file(s)
                ssef.write_excel_data_float( worksheet
                                      , final_iops_2nd, row1
                                      , col, 0, regular_10)
                                      
                
                ####################################
                # Calculating 
                #
                # Degradation is calculated with
                # respect to 1st file's data.
                ####################################
                degradation_1st_test = ssef.find_degradation(
                                                             file1_file_1st,
                                                             final_iops_1st
                                                             )
                
                degradation_2nd_test = ssef.find_degradation(
                                                             file1_file_2nd,
                                                             final_iops_2nd
                                                             )
                
                                  

                ##########################################
                # For statistical calculations
                #
                # It will calculate:
                #                  1.) Maximum
                #                  2.) Minimum
                #                  3.) Average
                #                  4.) Standard Deviation
                ##########################################
                
                
                #######################################
                # Writing Degradation and Staistics
                #
                #               1st Test 
                #######################################
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                            , degradation_1st_test, row2
                                            , col2, 0, regular_10)
                
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_1st_test, test_name_list[1])                  
                
                
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)
                
                
                ########################################
                # adding Series for 1st test
                # to the chart in "All Drives Baseline"
                ########################################
                chart.add_series({
                
                     #series name   
                    'name': str(test_name_list[1]), 
                    
                    # X-axis name
                    'categories': '='+str(ws_1st_name)+'!$A$' +str(row2+2)
                                    + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis values
                    'values': '='+str(ws_1st_name)+'!$' +str(svf.rank(col2+1))+
                                '$' +str(row2+2)+':$'+str(svf.rank(col2+1))+
                                '$'+str(row2+disk_no[0]+1),
                    
                    # Marker type            
                    'marker': {
                                'type': 'square',
                              }         
                    
                                    })                          
                    
                
                
                
                #######################################
                # Writing Degradation and Staistics
                #
                #               2nd Test 
                #######################################
                col2 +=1
                ssef.write_excel_data_float( worksheet
                                            , degradation_2nd_test, row2
                                            , col2, 0, regular_10)
                
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_2nd_test, test_name_list[2])                  
                        
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)
                
                # adding Series for 2nd test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                
                    #series name
                    'name': str(test_name_list[2]),
                    
                    # X-axis name
                    'categories': '='+ str(ws_1st_name)+'!$A$' 
                                     + str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis name    
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))+
                                 '$'+str(row2+disk_no[0]+1),
                    # Marker type             
                    'marker': {
                                'type': 'circle',
                              }   
                              
                                    })
                                
                
                ####################################
                # Writing IOMeter Error data 
                ####################################

                # Writing 2nd column
                col3 += 1  
                ssef.write_excel_data( worksheet
                                      , final_errors_1st, row3
                                      , col3, 0, regular_10)
                # Writing 3rd column
                col3 +=1                      
                ssef.write_excel_data_float( worksheet
                                          , final_errors_2nd, row3
                                          , col3, 0, regular_10)
                
                ############################
                #       FOR LOOP END!      #
                #         BASELINE         #
                ############################
                             
                             
                             
        #########################################
        # Finding Average of all Baseline
        # and writing in Worksheet
        #
        #           1st Test 
        #########################################
        
        # Sum divided by total number of files, 1st Test
        avg_1st = (null_array_1st/(sf_len)).tolist()

        # Heading for Average, 1st Test
        avg_1st_text = ['Overall Baseline: \n'+str(just_test_name[0]), 'Calculated Average']

        # Writing IOps Error, 1st Test
        col = 1 + (sf_len*2) + 1
        row=0
        row+=4

        ssef.write_excel_data( worksheet
                              , avg_1st_text, row
                              , col, 0, regular_10)
         
        row1 = row + 2        
        ssef.write_excel_data_float( worksheet
                                    , avg_1st, row1
                                    , col, 0, regular_10)
                                  
                
        #########################################
        # Finding Average of all Baseline
        # and writing in Worksheet
        #
        #           2nd Test 
        #########################################
        
        # Sum divided by total number of files, 2nd Test
        avg_2nd = (null_array_2nd/(sf_len)).tolist()
        
        # Heading for Average, 2nd Test
        avg_2nd_text = ['Overall Baseline: \n'+str(just_test_name[1]),
                        'Calculated Average']
        
        # Writing IOps Error, 1st Test
        col = 2 + (sf_len*2) + 1                            
        ssef.write_excel_data( worksheet
                              , avg_2nd_text, row
                              , col, 0, regular_10)
                 
        ssef.write_excel_data_float( worksheet
                                    , avg_2nd, row1
                                    , col, 0, regular_10)

        # Insert Chart in B2 Cell                          
        worksheet1.insert_chart('B2', chart)
        
        return [merge_format, bold_14, bold_12, bold_10, bold_10_l, regular_10, regular_10_l,
                workbook, avg_1st, avg_1st_text, avg_2nd, avg_2nd_text]
        
        
                    #############################
                    #                           #
#####################      Op Baseline END      #########################
                    #                           #
                    #############################







        
    #################################################
    # 1.) This function will take Average of IOps
    # from Baseline files and calculate
    # degradation with respect to it.
    #
    # 2.) It will also create and write
    # IOps, % Degradation & Error data
    # in "ws_1st_name" worksheet. 
    # Usually it is "Summary"    
    #
    # 3.) Now using this data. This function will
    # plot two charts. 
    #               a.) All Drives Vibe(ws_2nd_name)
    #               b.) High/Low/Average(ws_5th_name)
    ###################################################
    def generate_excel_performance(merge_format, bold_14, bold_12
                       , bold_10, bold_10_l, regular_10, regular_10_l, c_path
                       , sorted_files, file_dir_perf, workbook
                       , chassis_name, ws_1st_name, ws_2nd_name
                       , ws_5th_name, selector, avg_1st, avg_1st_text
                       , avg_2nd, avg_2nd_text, title):
        
        import os
        import xlsxwriter
        import sv_functions
        
        svf = sv_functions.SV_Functions

        
        ###################################
        #  Importing from other Directory
        ###################################
        import sys
        import os
        import numpy as np

        
        ##################################
        # Inserting "SSE" path to import 
        # and use some of its functions
        ##################################        
        os.chdir('..')
        c_path = os.getcwd()
        sys.path.insert(0, r''+str(c_path)+'/SSE')

        import sse_functions
        ssef= sse_functions.SSE_Functions
        
        
        ##################################
        # Jumping back to "SV" directory
        # This is our original directory
        ##################################
        sys.path.insert(0, r''+str(c_path)+'/SV')                               
                
                

        # length/number of sorted baseline files
        sf_len = len(sorted_files) 
        

                                       
        #############################
        # Create Baseline Worksheet
        #############################
        worksheet = workbook.add_worksheet(''+str(ws_1st_name))
        
        # Set column width to 30 points
        worksheet.set_column('A:Z', 30)

        
        # Setting borders
        regular_10.set_border(1)
        bold_10.set_border(1)
        bold_10_l.set_border(1)
        regular_10_l.set_border(1)
        
        format1 = workbook.add_format({'bold':True, 'align': 'center'})
        format1.set_num_format('0.0')
        format1.set_border()
        
        #########################################
        # Merge 1st & 2nd rows, over columns.
        #########################################
        worksheet.merge_range('A1:Z1', 'VIBRATION TEST RESULTS', merge_format)
        worksheet.merge_range('A2:Z2', ''+str(title), merge_format)
        
        
        #########################################
        # Including Title for IOps Result,
        # "col" & "row" indicates position 
        # in the Excel sheet.
        #########################################
        col = 0
        row = 3
        worksheet.write(row, col, 'IOps Results', bold_12)
        
        
        ################################
        # Writing "All Drives Baseline"                          
        ################################                                              
        worksheet1 = workbook.add_worksheet(''+str(ws_2nd_name))                          
        

        #########################################
        # Chart Settings
        #########################################        
        # chart type
        chart = workbook.add_chart({'type': 'line'})
        
        # chart scale
        chart.set_size({'x_scale': 3, 'y_scale': 2.5}) 
        
        # chart title
        chart.set_title({
                         'name':''+str(chassis_name)
                         +' Testing'\
                           '\nwith '+str(title)+
                            '\nIOMeter Test Results by Drive '
                        })
        
        # chart y axis range and Axis name                        
        chart.set_y_axis({
                            'name': 'IOps % Degradation',
                            'min': -15,
                            'max': 15
                         })
                         
        # chart x axis range, label position, gridlines, and Axis name                        
        chart.set_x_axis({
                            'label_position': 'low',
                            'position_axis': 'on_tick',
                            'min': 1,
                            'major_gridlines': 
                                            {
                                            'visible': True,
                                            'line': {'width': 0.5}
                                            }
                         })


        ###################################
        # This loop will read and extract  
        # Baseline file one by one. 
        #
        # It will write IOps data, 
        # Degradation, IOMeter Errors for 
        # both tests.
        ###################################
        
        #########################################
        # Performance file(s) should be in Y,X,Z 
        # order. Qualify each axis and making a 
        # list in required order.
        #########################################
        x_axis_files = svf.filter(sorted_files, 'X', 0)
        y_axis_files = svf.filter(sorted_files, 'Y', 0)
        z_axis_files = svf.filter(sorted_files, 'Z', 0)
        final_perf_files = y_axis_files + x_axis_files + z_axis_files
        
        # reading only Swept Sine(SS) file(s)
        ss_files = svf.filter(final_perf_files, 'SS', 0)
        #print('SS:' +str(len(ss_files)))
        
        # reading only Op Shock file(s)
        op_shock_files1 = svf.filter(final_perf_files, 'Op Shock', 0)
        op_shock_files2 = svf.filter(final_perf_files, 'OP Shock', 0)
        op_shock_files3 = svf.filter(final_perf_files, 'op Shock', 0)
        op_shock_files4 = svf.filter(final_perf_files, 'OP shock', 0)
        op_shock_files5 = svf.filter(final_perf_files, 'op shock', 0)
        op_shock_files = op_shock_files1 + op_shock_files2 + op_shock_files3 + op_shock_files4 + op_shock_files5
        #print('Op shock: '+str(len(op_shock_files)))
        
        # reading only Office Vibe, they needs to in group of two
        # Each needs to be treated differently
        office_vibe_files1 = svf.filter(final_perf_files, 'Office', 0)
        office_vibe_files2 = svf.filter(final_perf_files, 'office', 0)
        office_vibe_files = office_vibe_files1 + office_vibe_files2
        #print('Office Vibe:'+str(len(office_vibe_files)))
        
        # As 1st & 2nd test are in Individual files.
        # Making a group of both file.
        office_vibe_nested_list = []
        for i in range(0, int(len(office_vibe_files)/2)):
            
            temp = [office_vibe_files[2*i]] + [office_vibe_files[(2*i)+1]]
            
            office_vibe_nested_list.append(temp)
        
        
        #Final list of SS, Op shock & Office Vibe files in sequence
        perf_list = ss_files + office_vibe_nested_list + op_shock_files
        #print('Total len:'+str(len(perf_list)))
        
        # size of file list
        sf_len = len(perf_list)
                
        
        for sf in range(sf_len):
            
                        
            #############################
            # As Office Vibe file have 
            # each tests in individual 
            # files. So each step(sf)
            # will read and extract
            # data from both files.
            #############################
            if sf >= len(ss_files) and sf < (len(ss_files) + len(op_shock_files) ): # Office Vibe files
                
                #single list if group of two
                # Office vibe files
                performance_files = perf_list[sf] 
                
                
                # 1st Test
                [final_iops_1st, final_errors_1st
                , final_drives, test_name_list_1st
                , final_time_stamps_1st
                , just_test_name_1st] = svf.extract_office_vibe_single_file(file_dir_perf,
                                                                            performance_files[0],
                                                                            selector)
                
                # 2nd Test
                [final_iops_2nd, final_errors_2nd
                , final_drives, test_name_list_2nd
                , final_time_stamps_2nd
                , just_test_name_2nd] = svf.extract_office_vibe_single_file(file_dir_perf, 
                                                                            performance_files[1],
                                                                            selector)
                
                
                just_test_name = just_test_name_1st + just_test_name_2nd
                        
                        
                # Joining/Modifying List so that we can directly write
                # it in 1st Column
                # 1st column,IOps
                col_1st = ['Time Stamp'] + [test_name_list_1st[0]] + final_drives
                
                # 2nd column,IOps
                test_name1 = test_name_list_1st[1]
                col_2nd = [final_time_stamps_1st[0]] + [test_name1] 
                
                # 3rd column,IOps
                test_name2 = test_name_list_2nd[1]
                col_3rd = [final_time_stamps_2nd[0]] + [test_name2]
    
    
    

            ##############################
            # Rest of the file(s) has 
            # two tests in each individual
            # file(assumed).
            ##############################
            else : # Swept sine & Op shock files
            
                [final_iops_1st, final_errors_1st
                , final_iops_2nd, final_errors_2nd
                , final_drives, test_name_list
                , final_time_stamps, disk_no
                , just_test_name] = svf.extract_baseline_data(file_dir_perf, 
                                                              perf_list[sf], 
                                                              selector)
                                                              
                                                              
                ########################################################
                # Joining/Modifying List so that we can directly write
                # it in 1st Column
                # 1st column,IOps
                ########################################################
                
                col_1st = [final_time_stamps[0]] + [test_name_list[0]] + final_drives
                
                # 2nd column,IOps
                test_name1 = test_name_list[1]
                col_2nd = [final_time_stamps[1]] + [test_name1] 
                
                # 3rd column,IOps
                test_name2 = test_name_list[2]
                col_3rd = [final_time_stamps[2]] + [test_name2]
            
            ###########################
            # Data Extraction done!
            ###########################
        
        
        
            ##############################
            # Manipulate & write extracted
            # data.
            ##############################
            # We need to write 1st column
            # just once. So if it is a 
            # 1st file. Write 1st, 2nd, 
            # and 3rd column.
            ##############################
            if sf == 0:
            
                # store data of 1st file to find degradation value later
                file1_file_1st = final_iops_1st 
                file1_file_2nd = final_iops_2nd
                
                
                ######################################################
                # creating an empty list of zeros, which will finally
                # hold row-wise summation of All .csv files.
                # Which is used to calculate the Average
                ######################################################
                
                # 1st Test
                null_array_1st = np.zeros(int(disk_no[0])+1)
                null_array_1st = null_array_1st.tolist() 
                
                # 2nd Test
                null_array_2nd = null_array_1st
                
                
                #########################################
                # Zeros array to calculate average
                #########################################
                
                # Summing up 1st test data
                null_array_1st = np.sum([final_iops_1st, null_array_1st], axis = 0)

                # Summing up 2nd test data
                null_array_2nd = np.sum([final_iops_2nd, null_array_2nd], axis = 0)
                      
                
                                
                #############################################                    
                # Writing 1st column, IOps, Left justified     
                # Originally row = 3, for IOps title.
                # Now row =4, to write it from next 
                # row, col = 0 (1st column)
                #############################################
                
                #####################
                # Drives
                #####################
                row+= 1 
                ssef.write_excel_data( worksheet
                                      , col_1st, row
                                      , col, 0, regular_10_l)


                
                
                #####################
                # Baseline Average
                #####################
                
                # Average of 1st test from Baseline 
                col+=1
                row1 = row + 2
                
                # used for Plotting in High/Low/Average 
                td_row = row + 1 
                
                # Title of 1st Test, from Baseline data
                ssef.write_excel_data( worksheet
                                      , avg_1st_text, row
                                      , col, 0, regular_10)
                
                # Average IOps of 1st Test, from Baseline data
                ssef.write_excel_data_float( worksheet
                                            , avg_1st, row1
                                            , col, 0, regular_10)
                        
                               
                # Title of 2nd Test, from Baseline data 
                col+=1
                ssef.write_excel_data( worksheet
                                      , avg_2nd_text, row
                                      , col, 0, regular_10)

                # Average IOps of 2nd Test, from Baseline data
                ssef.write_excel_data_float( worksheet
                                            , avg_2nd, row1
                                            , col, 0, regular_10)
                        
                
                                
                                
                #####################
                #  PERFORMANCE DATA
                #####################

                # Writing 2nd column, center justified
                col+=1  # Moving to next Column
               
               
                # Test 1, IOps title (Time stamp & Test Description )
                # this function will write given list 
                # of strings row-wise
                ssef.write_excel_data( worksheet
                                      , col_2nd, row
                                      , col, 0, regular_10)
 
                
                # Test 1, IOps data
                # this function will write given list 
                # of numbers, row-wise                      
                ssef.write_excel_data_float( worksheet
                                          , final_iops_1st, row1
                                          , col, 0, regular_10)

                       
                       
                # Writing 3rd column
                # Test 2, IOps title (Time stamp & Test Description )
                col+=1                      
                ssef.write_excel_data( worksheet
                                      , col_3rd, row
                                      , col, 0, regular_10)
                # Test 2, IOps data
                ssef.write_excel_data_float( worksheet
                                            , final_iops_2nd, row1
                                            , col, 0, regular_10)
                
            
                ####################################
                # Writing degradation data 
                ####################################
                # calculate next row to write data
                #####################################
                
                ###################################################
                #
                # starting row (row2)=   row1(staring IOps row) 
                #                       + (number of disks + system) 
                #                       + 4 Blank Rows  
                #
                ####################################################
                row2 = row1 + (disk_no[0]+1) + 4 # 4 Blank Rows
                col2 = 0 
                
                # Degradation title
                worksheet.write(row2, col2, 'IOps, % degradation, compared to overall baseline tests', bold_12)
                
                # Writing 1st column, Degradation 
                # Drives
                row2 += 1
                ssef.write_excel_data( worksheet
                                      , col_1st[2:], row2
                                      , col2, 0, regular_10_l)
                
                
                #####################################
                # As degradation is calculated with 
                # respect to 1st file's values
                # leave 2nd and 3rd column blank.
                #
                # Write Statistical titles in 3rd column
                #########################################    
                col2 += 2
                row2_new = row2 + (disk_no[0]+1)
                
                start_col = col2
                # To write Statistical data
                stat_list = ['High', 'Low', 'Average', '1 σ']
                
                ssef.write_excel_data( worksheet
                                      , stat_list, row2_new
                                      , col2, 0, bold_10_l)
                
                
                ####################################
                # Calculating 
                #
                # Degradation is calculated with
                # respect to 1st file's data.
                ####################################
                
                ################################
                # Comparing test with baseline 
                # so that Degradation is calculated
                # with correct Baseline average 
                # Test.
                ################################
                import re
                
                # exracting just numbers to compare
                jtn_1st = re.findall(r'\d+', just_test_name[0])
                jtn_2nd = re.findall(r'\d+', just_test_name[1])
                
                
                at_1st = re.findall(r'\d+', avg_1st_text[0])
                at_2nd = re.findall(r'\d+', avg_2nd_text[0])
                
                #print(at_1st, at_2nd)
                
                # 1st Test
                # if name of 1st same as 1st test's name from Baseline
                if int(jtn_1st[0]) == int(at_1st[0]):  
                    degradation_1st_test = ssef.find_degradation(avg_1st, final_iops_1st)
                
                # if name of 1st same as 2nd test's average name from Baseline
                elif int(jtn_1st[0]) == int(at_2nd[0]):
                    degradation_1st_test = ssef.find_degradation(avg_2nd, final_iops_1st)
                
                # else notify User about unmatch
                else:
                    print('\nTest name of Baseline and Performance file does not match!'\
                          '\nPlease make sure it matches, to generate a report.')

                      
                # 2nd test          
                # if name of 2nd test same as 1st test's name from Baseline
                if int(jtn_2nd[0]) == int(at_1st[0]):
                    degradation_2nd_test = ssef.find_degradation(avg_1st, final_iops_2nd)
                
                # if name of 2nd test same as 2nd test's average name from Baseline
                elif int(jtn_2nd[0]) == int(at_2nd[0]):
                    degradation_2nd_test = ssef.find_degradation(avg_2nd, final_iops_2nd)
                
                # else notify User about unmatch
                else:
                    print('\nTest name of Baseline and Performance file does not match!'\
                          '\nPlease make sure it matches, to generate a report.')
                          
                
                # Write degradation of 1st test
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                      , degradation_1st_test, row2
                                      , col2, 0, regular_10)
                                      
                # Perform Conditional formatting on given "degradation_1st_test" list
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_1st_test, jtn_1st[0])                  
                    
                
                ##########################################
                # For statistical calculations
                #
                # It will calculate:
                #                  1.) Maximum
                #                  2.) Minimum
                #                  3.) Average
                #                  4.) Standard Deviation
                ##########################################
                
                # For 1st test
                # stat_1st_test = [round(float(np.max(degradation_1st_test[1:])),1)] + [round(float(np.min(degradation_1st_test[1:])),1)] + [round(float(np.mean(degradation_1st_test[1:])),1)]+ [round(float(np.std(degradation_1st_test[1:])),1)]
                                       
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)

                
                # adding Series for 1st test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                
                    #series name
                    'name': str(test_name1),
                    
                    # X-axis name
                    'categories': '='+str(ws_1st_name)+'!$A$' 
                                     +str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    # Y-axis values                        
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))
                                 +'$'+str(row2+disk_no[0]+1),
                    
                    # Marker type
                    'marker': {
                                'type': 'square',
                              }         
                                 })                          
                
                
                #########################
                # 2nd Test 
                #########################
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                      , degradation_2nd_test, row2
                                      , col2, 0, regular_10)
                
                # Perform Conditional formatting on given "degradation_2nd_test" list
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_2nd_test, jtn_2nd[0])                  
                 
                
                               
                # For 2nd test, Staistics                  
                # stat_2nd_test = [round(float(np.max(degradation_2nd_test[1:])),1)] + [round(float(np.min(degradation_2nd_test[1:])),1)] + [round(float(np.mean(degradation_2nd_test[1:])),1)]+ [round(float(np.std(degradation_2nd_test[1:])),1)]         
                
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)

                
                # adding Series for 2nd test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                    
                    #series name
                    'name': str(test_name2),
                    
                    # X-axis name
                    'categories': '='+ str(ws_1st_name)+'!$A$' 
                                     + str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                     
                    # Y-axis values
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))
                                 +'$'+str(row2+disk_no[0]+1),
                    
                    # Marker type
                    'marker': {
                                'type': 'circle',
                              }   
                              
                                })
                
                
                ####################################
                # Writing IOMeter Error data 
                ####################################
                # calculate next row to write data
                #####################################
                
                ###################################################
                #
                # starting row (row3)=   row2(staring Degradation row) 
                #                       + (number of disks + system) 
                #                       + 6 Blank Rows  
                #
                ####################################################
                row3 = row2 + (disk_no[0]+1) + 6 
               
                
                # Error Title
                col3 = 0 
                worksheet.write(row3, col3, 'IOMeter Errors', bold_12)
                
                # Writing 1st column, Degradation, Drives #     
                row3 += 1
                ssef.write_excel_data( worksheet
                                      , col_1st[2:], row3
                                      , col3, 0, regular_10_l)
                
                
                # Writing 2nd column, 1st Test Errors
                col3 = 2
                col3 += 1  
                      
                ssef.write_excel_data( worksheet
                                      , final_errors_1st, row3
                                      , col3, 0, regular_10)
                
                # Writing 3rd column, 2nd Test Errors
     
                col3+=1                      
                ssef.write_excel_data_float( worksheet
                                          , final_errors_2nd, row3
                                          , col3, 0, regular_10)

            
            
            
            ##############################
            #     PERFORMANCE DATA
            ##############################
            # We need to write 1st column
            # just once. So if it is a 
            # 1st file. Write 1st, 2nd, 
            # and 3rd column.
            #
            # Else just write 2nd and 3rd 
            # column from 2nd and rest of 
            # files.
            ##############################
                                         
            else:
            
                #########################################
                # Null array to calculate average
                #########################################
                
                # Summing up 1st test data
                null_array_1st = np.sum([final_iops_1st, null_array_1st], axis = 0)

                # Summing up 2nd test data
                null_array_2nd = np.sum([final_iops_2nd, null_array_2nd], axis = 0)
         
         
                # this function will write given list 
                # of strings row-wise
                row1 = row + 2 
                
                
                ##################################
                # 2n + 3 series, It will generate
                # odd numbers.
                #
                # In our case, as sf starts from 1,
                # we will get:
                #              5,7,9...(2n+1)
                #
                # Which is 6th,8th...(2n+3) columns
                # from Excel point of view. As it 
                # starts numbering from 1.
                #
                # So in Excel worksheet, even columns 
                # will have 1st test data.
                #####################################
                col = 3 + (sf*2)         
                
                # 1st Test, title (Time stamp & Test Description)
                ssef.write_excel_data( worksheet
                                      , col_2nd, row
                                      , col, 0, regular_10)
                
                
                # this function will write given list 
                # of numbers, row-wise 
                # Writing 2nd column
                
                # 1st Test, IOps data
                ssef.write_excel_data_float( worksheet
                                              , final_iops_1st, row1
                                              , col, 0, regular_10)
                         
                
                ##################################
                # 2n + 4 series, It will generate
                # even numbers.
                #
                # In our case, as sf starts from 1,
                # we will get:
                #              6,8,10,...,(2n+5)
                #
                # Which is 7th, 9th, ...,(2n+5) columns
                # from Excel point of view. As it 
                # starts numbering from 1.
                #
                # So in Excel worksheet, odd columns 
                # will have 2nd test data from Same 
                # file.
                #####################################
                col = 4 + (sf*2)                       
                
                # Writing 3rd column, 
                # 2nd Test, title (Time stamp & Test Description)
                ssef.write_excel_data( worksheet
                                      , col_3rd, row
                                      , col, 0, regular_10)
                      
                # 2nd Test, IOps data
                ssef.write_excel_data_float( worksheet
                                      , final_iops_2nd, row1
                                      , col, 0, regular_10)
                                      
                
                ####################################
                # Calculating 
                #
                # Degradation is calculated with
                # respect to Average of that test.
                ####################################
                
                ################################
                # Comparing test with baseline 
                # so that Degradation is calculated
                # with correct Baseline average 
                # Test.
                ################################
                
                jtn_1st = re.findall(r'\d+', just_test_name[0])
                jtn_2nd = re.findall(r'\d+', just_test_name[1])
                
                
                at_1st = re.findall(r'\d+', avg_1st_text[0])
                at_2nd = re.findall(r'\d+', avg_2nd_text[0])
               
                
                # 1st Test
                # if name of 1st same as 1st test's name from Baseline
                if int(jtn_1st[0]) == int(at_1st[0]):
                    degradation_1st_test = ssef.find_degradation(avg_1st, final_iops_1st)
                
                # if name of 1st same as 2nd test's average name from Baseline
                elif int(jtn_1st[0]) == int(at_2nd[0]):
                    degradation_1st_test = ssef.find_degradation(avg_2nd, final_iops_1st)
                
                # else notify User about unmatch
                else:
                    print('\nTest name of Baseline and Performance file does not match!'\
                          '\nPlease make sure it matches, to generate a report.')
                      


                      
                # 2nd test   
                # if name of 2nd test same as 1st test's name from Baseline                
                if int(jtn_2nd[0]) == int(at_1st[0]):
                    degradation_2nd_test = ssef.find_degradation(avg_1st, final_iops_2nd)
                
                # if name of 2nd test same as 2nd test's average name from Baseline
                elif int(jtn_2nd[0]) == int(at_2nd[0]):
                    degradation_2nd_test = ssef.find_degradation(avg_2nd, final_iops_2nd)
                
                 # else notify User about unmatch
                else:
                    print('\nTest name of Baseline and Performance file does not match!'\
                          '\nPlease make sure it matches, to generate a report.')
                
                                  

                ##########################################
                # For statistical calculations
                #
                # It will calculate:
                #                  1.) Maximum
                #                  2.) Minimum
                #                  3.) Average
                #                  4.) Standard Deviation
                ##########################################
               
                #######################################
                # Writing Degradation and Staistics
                #
                #               1st Test 
                #######################################
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                      , degradation_1st_test, row2
                                      , col2, 0, regular_10)
                                      
                # Perform Conditional formatting on given "degradation_1st_test" list
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_1st_test, jtn_1st[0])                  
                                     
                    
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)
                
                
                # adding Series for 1st test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                    
                    #series name
                    'name': str(test_name1),
                    
                    # X-axis
                    'categories': '='+str(ws_1st_name)+'!$A$' 
                                     +str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis    
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))
                                 +'$'+str(row2+disk_no[0]+1),
                    
                    # marker type
                    'marker': {
                                'type': 'square',
                              }         
                                 
                                 })                          
                    
                
                
                
                #######################################
                # Writing Degradation and Staistics
                #
                #               2nd Test 
                #######################################
                col2 +=1
                
                # degradation data
                ssef.write_excel_data_float( worksheet
                                      , degradation_2nd_test, row2
                                      , col2, 0, regular_10)
                                      
                # Perform Conditional formatting on given "degradation_2nd_test" list
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_2nd_test, jtn_2nd[0])                  
                 
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)

                
                
                # adding Series for 2nd test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                    
                    #series name
                    'name': str(test_name2),
                        
                    # X-axis
                    'categories': '='+str(ws_1st_name)+'!$A$' 
                                     +str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis    
                    'values': '='+str(ws_1st_name)+'!$'
                                 +str(svf.rank(col2+1))+'$' +str(row2+2)
                                 +':$'+str(svf.rank(col2+1))+'$'+str(row2+disk_no[0]+1),
                    
                    # marker type
                    'marker': {
                                'type': 'circle',
                              }   
                              
                                })
                                
                
                ####################################
                # Writing IOMeter Error data 
                ####################################

                # Writing 2nd column
                col3 += 1  
                ssef.write_excel_data( worksheet
                                      , final_errors_1st, row3
                                      , col3, 0, regular_10)
                # Writing 3rd column
                col3 +=1                      
                ssef.write_excel_data_float( worksheet
                                          , final_errors_2nd, row3
                                          , col3, 0, regular_10)
                
                ############################
                #       FOR LOOP END!      #
                ############################
        
                     
        # Inserting "Line chart" after 
        # Loop ends at B2
        worksheet1.insert_chart('B2', chart)
        
 

 
        
        #########################
        # Creating 
        # Hi, Lo, Avg Chart in 
        # a given Workbook.
        #########################
        worksheet = workbook.add_worksheet(''+str(ws_5th_name))

        chart = workbook.add_chart({'type': 'stock'})
        
        chart.set_size({'x_scale': 3 , 'y_scale': 2.5}) 
        
        # chart title
        chart.set_title({'name':''+str(chassis_name)
                                  +' Testing'\
                                   '\nwith '+str(title)+
                                   '\nIOMeter Test Results, High/Low/Average'})
        
        # chart y axis range and Axis name                        
        chart.set_y_axis({'name': 'IOps % Degradation',
                         'min': -20, 'max': 20})
                         
        # chart x axis range, label position, gridlines, and Axis name                        
        chart.set_x_axis({
                            'label_position': 'low'
                          , 'position_axis': 'on_tick'
                          , 'min': 1
                          , 'major_gridlines': {
                                                'visible': True,
                                                'line': {'width': 0.5}
                                                }
                         })

                         
        ####################
        # Adding Series
        ####################
        # High
        ###################
        chart.add_series({
                            # shape
                            'line': 'o',
                            
                            # name
                            'name': 'High',
                            
                            # X-axis
                            'categories': '='+str(ws_1st_name)
                                             +'!$'+str(svf.rank(start_col+2)) 
                                             +str(td_row+1)
                                             + ':$'+str(svf.rank(col2+1))
                                             +'$'+str(td_row+1),
                            # Y-axis    
                            'values': '='+str(ws_1st_name)
                                         +'!$'+str(svf.rank(start_col+2)) 
                                         +str(row2_new+1)
                                         + ':$'+str(svf.rank(col2+1))
                                         +'$'+str(row2_new+1),
                     
                            'line':   {'none': True},    
                            #'line':       {'color': 'red'},
                            'marker': {
                                        'type': 'diamond',
                                        'border': {'color': 'red'},
                                        'fill':   {'color': 'red'},
                                    }

                        })
                        
                        
        #################
        # Low
        #################
        chart.add_series({
                            # series name
                            'name': 'Low',
                            
                            # X-axis
                            'categories': '='+str(ws_1st_name)
                                             +'!$'+str(svf.rank(start_col+2)) 
                                             +str(td_row+1)
                                             + ':$'+str(svf.rank(col2+1))
                                             +'$'+str(td_row+1),
                            # Y-axis    
                            'values':     '='+str(ws_1st_name)+'!$'+str(svf.rank(start_col+2)) 
                                            +str(row2_new+2)
                                            + ':$'+str(svf.rank(col2+1))+'$'+str(row2_new+2),
                     
                                
                            'line':   {'none': True},
                           
                            # Marker properties
                            'marker': {
                                        'type': 'square',
                                        'border': {'color': 'blue'},
                                        'fill': {'color':'blue'}
                                      }
                          })

        #################
        # Average
        #################
        chart.add_series({
                            # series name
                            'name': 'Average',
                            
                            # X-axis
                            'categories': '='+str(ws_1st_name)+'!$'
                                             +str(svf.rank(start_col+2)) 
                                             +str(td_row+1)
                                             + ':$'+str(svf.rank(col2+1))
                                             +'$'+str(td_row+1),
                            # Y-axis    
                            'values': '='+str(ws_1st_name)+'!$'
                                         +str(svf.rank(start_col+2)) 
                                         +str(row2_new+3)
                                         + ':$'+str(svf.rank(col2+1))+
                                         '$'+str(row2_new+3),
                            # line property
                            'line':   {'none': True},
                            
                            # marker properties
                            'marker': {
                                        'type': 'square',
                                        'border': {'color': 'green'},
                                        'fill': {'color':'green'}
                                      }
                        })
                        
                        
        
        
        # Insert High/Low/Average chart at B2
        worksheet.insert_chart('B2', chart)
        
        return workbook
        
                    #############################
                    #                           #
#####################     Op PERFORMANCE END    #########################
                    #                           #
                    #############################

        
        
        
        
    #################################
    #           NON-OP              #
    #################################    
    
    #############################################
    # Function will extract data from 
    # given Baseline files from main 
    # file. Manipulate it and write it
    # in worksheet "ws_1st_name".
    #
    # This function is similar to
    # "generate_excel_baseline". Only difference
    # is it will compare degradation with respect
    # to Average of IOps instead of 1st file.
    #
    # Using this data, it will plot and
    # add chart to "ws_2nd_name" worksheet
    ##############################################
    def generate_excel_non_op_baseline(c_path, sorted_files, file_dir_perf, workbook
                              , chassis_name, ws_1st_name, ws_2nd_name, selector, title):
        
        
        import os
        import xlsxwriter
        
        import sv_functions
        svf = sv_functions.SV_Functions

        
        ###################################
        #  Importing from other Directory
        ###################################
        import sys
        import os
        import numpy as np

        ##################################
        # Inserting "SSE" path to import 
        # and use some of its functions
        ##################################        
        os.chdir('..')
        c_path = os.getcwd()
        sys.path.insert(0, r''+str(c_path)+'/SSE')

        import sse_functions
        ssef= sse_functions.SSE_Functions

        ##################################
        # Jumping back to "SV" directory
        # This is our original directory
        ##################################
        sys.path.insert(0, r''+str(c_path)+'/SV')                               
                
        
        
        #########################################
        # This loop is used just to calculate
        # average of all files
        #
        #########################################
        
        
                
        sf_len = len(sorted_files)
        for sf in range(sf_len):
            
            # initialize null_array_1st & null_array_2nd list only for 1st file
            if sf == 0:
                
                ######################################
                # This function will extract all data 
                # from Single .csv Baselie file.
                ######################################            
                [final_iops_1st, final_errors_1st
                , final_iops_2nd, final_errors_2nd
                , final_drives, test_name_list
                , final_time_stamps, disk_no
                , just_test_name] = svf.extract_baseline_data(file_dir_perf,
                                                              sorted_files[sf],
                                                              selector)
                                                             
                # Zeros array to calculate average
                # 1st Test
                null_array_1st = np.zeros(int(disk_no[0])+1)
                null_array_1st = null_array_1st.tolist() 
                
                # 2nd Test
                null_array_2nd = null_array_1st

                
                # Summing up 1st test data
                null_array_1st = np.sum([final_iops_1st, null_array_1st], axis = 0)

                # Summing up 2nd test data
                null_array_2nd = np.sum([final_iops_2nd, null_array_2nd], axis = 0)

            else:
                ######################################
                # This function will extract all data 
                # from Single .csv Baselie file.
                ######################################            
                [final_iops_1st, final_errors_1st
                , final_iops_2nd, final_errors_2nd
                , final_drives, test_name_list
                , final_time_stamps, disk_no
                , just_test_name] = svf.extract_baseline_data(file_dir_perf,
                                                              sorted_files[sf],
                                                              selector)
                                                             
                # Summing up 1st test data
                null_array_1st = np.sum([final_iops_1st, null_array_1st], axis = 0)

                # Summing up 2nd test data
                null_array_2nd = np.sum([final_iops_2nd, null_array_2nd], axis = 0)

        ################################
        # Average of IOps
        ################################
        
        # Sum divided by total number of files, 1st Test
        avg_1st = (null_array_1st/(sf_len)).tolist()
        
        # Sum divided by total number of files, 2nd Test
        avg_2nd = (null_array_2nd/(sf_len)).tolist()
        
        
        
        
        #############################
        # Create Baseline Worksheet
        #############################
        worksheet = workbook.add_worksheet(''+str(ws_1st_name))
        
        # Set column width to 30 points
        worksheet.set_column('A:Z', 30)

        
        ############################
        # Create & Formate 
        # "Baseline" worksheet in 
        # a Report.
        ############################
        
        # Create a format to use in the merged range.
        merge_format = workbook.add_format({
                                            'border': 1,
                                            'align': 'left',
                                            'valign': 'left',
                                            'fg_color': 'yellow',
                                            'bold': True,
                                            'font_name':'Arial',
                                            'font_size':14
                                           })

        
        # Add a bold format to use to highlight cells.
        bold_14 = workbook.add_format({
                                        'bold': True, 
                                        'font_name':'Arial',
                                        'font_size':14
                                      })


      
        bold_12 = workbook.add_format({ 
                                        'bold': True,
                                        'font_name':'Arial',
                                        'font_size':12
                                      })

                                      
        bold_10 = workbook.add_format({
                                        'bold': True, 
                                        'font_name':'Arial',
                                        'font_size':10,
                                        'align': 'center'
                                      })
                                      
        bold_10_l = workbook.add_format({
                                        'bold': True, 
                                        'font_name':'Arial',
                                        'font_size':10,
                                        'align': 'left'
                                      })
        
        
        regular_10 = workbook.add_format({
                                          'font_name':'Arial',
                                          'font_size':10,
                                          'align': 'center'
                                        })

        # Font: Arial, Font_size:10 & left Aligned
        regular_10_l = workbook.add_format({
                                            'font_name':'Arial',
                                            'font_size':10,
                                            'align': 'left'
                                           })
        
        # Set border for certain fonts
        regular_10.set_border(1)
        bold_10.set_border(1)
        bold_10_l.set_border(1)
        regular_10_l.set_border(1)
        
        format1 = workbook.add_format({'bold':True, 'align': 'center'})
        format1.set_num_format('0.0')
        format1.set_border()
        
        
        #########################################
        # Merge 1st & 2nd rows, over columns.
        #########################################
        worksheet.merge_range('A1:Z1', 'BASELINE TEST RESULTS', merge_format)
        worksheet.merge_range('A2:Z2', ''+str(title), merge_format)
        
        
        #########################################
        # Including Title for IOps Result,
        # "col" & "row" indicates position 
        # in the Excel sheet.
        #########################################
        col = 0
        row = 3
        worksheet.write(row, col, 'IOps Results', bold_12)
        
        
        ################################
        # Writing "All Drives Baseline"                          
        ################################                                              
        worksheet1 = workbook.add_worksheet(''+str(ws_2nd_name))                          
        

        #########################################
        # Chart Settings
        #########################################        
        
        # chart type
        chart = workbook.add_chart({'type': 'line'})
        
        # chart scale
        chart.set_size({'x_scale': 3, 'y_scale': 2.5}) 
        
        # chart title
        chart.set_title({
                         'name':''+str(chassis_name)
                         +'  Basline Testing'\
                         '\nwith ' +str(title)+
                         '\nIOMeter Test Results by Drive '
                        })
        
        
        # chart y axis range and Axis name                        
        chart.set_y_axis({
                           'name': 'IOps % Degradation',
                           'min': -15,
                           'max': 15
                         })
                         
                         
        # chart x axis range, label position, gridlines, and Axis name                        
        chart.set_x_axis({
                           'label_position': 'low',
                           'position_axis': 'on_tick',
                           'min': 1,
                           'major_gridlines': {
                                                'visible': True,
                                                'line': {'width': 0.5}
                                              }
                         })


                         
        ###################################
        # This loop will read and extract  
        # Baseline file one by one. 
        #
        # It will write IOps data, 
        # Degradation, IOMeter Errors for 
        # both tests.
        ###################################
        for sf in range(sf_len):
            
            ######################################
            # This function will extract all data 
            # from Single .csv Baselie file.
            ######################################            
            [final_iops_1st, final_errors_1st
            , final_iops_2nd, final_errors_2nd
            , final_drives, test_name_list
            , final_time_stamps, disk_no
            , just_test_name] = svf.extract_baseline_data(file_dir_perf,
                                                          sorted_files[sf],
                                                          selector)
            #print(len(final_iops_1st))
            ######################################                                                          
            # Joining/Modifying List so that we 
            # can directly write it in 1st Column
            ######################################
            
            # 1st column,IOps
            col_1st = [final_time_stamps[0]] + [test_name_list[0]] + final_drives
            
            # 2nd column,IOps
            col_2nd = [final_time_stamps[1]] + [test_name_list[1]] 
            
            # 3rd column,IOps
            col_3rd = [final_time_stamps[2]] + [test_name_list[2]]
            
            
            
            ##############################
            # We need to write 1st column
            # just once. So if it is a 
            # 1st file. Write 1st, 2nd, 
            # and 3rd column.
            ##############################
            if sf == 0:
            
                
                #############################################                    
                # Writing 1st column, IOps, Left justified     
                # Originally row = 3, for IOps title.
                # Now row =4, to write it from next 
                # row, col = 0 (1st column)
                #############################################
                
                row+= 1 
                ssef.write_excel_data( worksheet
                                      , col_1st, row
                                      , col, 0, regular_10_l)
                
                
                # Writing 2nd column, center justified
                col+=1  # Moving to next Column
                row1 = row + 2 
                
                #############################################                 
                # this function will write given list 
                # of strings row-wise
                #############################################                
                ssef.write_excel_data( worksheet, col_2nd, row, 
                                       col, 0, regular_10)
                
                
                #######################################
                # this function will write given list 
                # of numbers, row-wise                      
                #######################################
                ssef.write_excel_data_float( worksheet
                                           , final_iops_1st, row1
                                           , col, 0, regular_10)
                             
                             
                # Writing 3rd column
                col+=1                      
                ssef.write_excel_data( worksheet, col_3rd, row,
                                       col, 0, regular_10)
                                      
                ssef.write_excel_data_float( worksheet, final_iops_2nd, 
                                             row1, col, 0, regular_10)
            
            
            
            
                ####################################
                # Writing degradation data 
                ####################################
                # calculate next row to write data
                #####################################
                
                ###################################################
                #
                # starting row (row2)=   row1(staring IOps row) 
                #                       + (number of disks + system) 
                #                       + 4 Blank Rows  
                #
                ####################################################
                row2 = row1 + (disk_no[0]+1) + 4 
                col2 = 0 # 1st column 
                
                # Degradation title
                worksheet.write(row2, col2, 'IOps, % degradation, compared to Average baseline tests', bold_12)
                
                
                # Writing 1st column, Degradation     
                row2 += 1
                ssef.write_excel_data( worksheet, 
                                       col_1st[2:],
                                       row2, 
                                       col2, 
                                       0, 
                                       regular_10_l)
            
                #####################################
                # As degradation is calculated with 
                # respect to 1st file's values
                # leave 2nd and 3rd column blank.
                #
                # Write Statistical titles in 3rd column
                #########################################
                #col2 += 2
                row2_new = row2 + (disk_no[0]+1)
                
                # To write Statistical data titles
                stat_list = ['High', 'Low', 'Average', '1 σ']
                
                # Writing stat list
                ssef.write_excel_data( worksheet
                                      , stat_list, row2_new
                                      , col2, 0, bold_10_l)
                
                
                ####################################
                # Calculating 
                #
                # Degradation is calculated with
                # respect to 1st file's data.
                ####################################
                degradation_1st_test = ssef.find_degradation(
                                                             avg_1st,
                                                             final_iops_1st
                                                             )
                
                degradation_2nd_test = ssef.find_degradation(
                                                             avg_2nd,
                                                             final_iops_2nd
                                                             )
                
                                  

                ##########################################
                # For statistical calculations
                #
                # It will calculate:
                #                  1.) Maximum
                #                  2.) Minimum
                #                  3.) Average
                #                  4.) Standard Deviation
                ##########################################
                
                #######################################
                # Writing Degradation and Staistics
                #
                #               1st Test 
                #######################################
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                            , degradation_1st_test, row2
                                            , col2, 0, regular_10)
                
                # conditional formatting
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_1st_test, test_name_list[1])                  
                
                
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)                            
                                      
                                            
                ########################################
                # adding Series for 1st test
                # to the chart in "All Drives Baseline"
                ########################################
                chart.add_series({
                
                     #series name   
                    'name': str(test_name_list[1]), 
                    
                    # X-axis name
                    'categories': '='+str(ws_1st_name)+'!$A$' +str(row2+2)
                                    + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis values
                    'values': '='+str(ws_1st_name)+'!$' +str(svf.rank(col2+1))+
                                '$' +str(row2+2)+':$'+str(svf.rank(col2+1))+
                                '$'+str(row2+disk_no[0]+1),
                    
                    # Marker type            
                    'marker': {
                                'type': 'square',
                              }         
                    
                                    })                          
                    
                
                
                
                #######################################
                # Writing Degradation and Staistics
                #
                #               2nd Test 
                #######################################
                col2 +=1
                ssef.write_excel_data_float( worksheet
                                            , degradation_2nd_test, row2
                                            , col2, 0, regular_10)
                
                # conditional formatting
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_2nd_test, test_name_list[2])                  
                    
                
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)
               
                
                # adding Series for 2nd test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                
                    #series name
                    'name': str(test_name_list[2]),
                    
                    # X-axis name
                    'categories': '='+ str(ws_1st_name)+'!$A$' 
                                     + str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis name    
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))+
                                 '$'+str(row2+disk_no[0]+1),
                    # Marker type             
                    'marker': {
                                'type': 'circle',
                              }   
                              
                                    })
                                
                
                
                ####################################
                # Writing IOMeter Error data 
                ####################################
                # calculate next row to write data
                #####################################
                
                ###################################################
                #
                # starting row (row3)=   row2(staring Degradation row) 
                #                       + (number of disks + system) 
                #                       + 6 Blank Rows  
                #
                ####################################################
                row3 = row2 + (disk_no[0]+1) + 6  
                
                # Error Title
                col3 = 0 
                worksheet.write(row3, col3, 'IOMeter Errors', bold_12)
                
                # Writing 1st column, Drives #     
                row3 += 1
                ssef.write_excel_data( worksheet
                                      , col_1st[2:], row3
                                      , col3, 0, regular_10_l)
                
                # Writing 2nd column, 1st Test Errors
                col3+=1  
                      
                ssef.write_excel_data( worksheet
                                      , final_errors_1st, row3
                                      , col3, 0, regular_10)
                
                # Writing 3rd column, 2nd Test Errors                
                col3+=1                      
                ssef.write_excel_data_float( worksheet
                                          , final_errors_2nd, row3
                                          , col3, 0, regular_10)

            
            
            
            
            
            ##############################
            # We need to write 1st column
            # just once. So if it is a 
            # 1st file. Write 1st, 2nd, 
            # and 3rd column.
            #
            # Else just write 2nd and 3rd 
            # column from 2nd and rest of 
            # files.
            ##############################
                                         
            else:
                        
                              
                ##################################
                # 2n + 1 series, It will generate
                # odd numbers.
                #
                # In our case, as sf starts from 1,
                # we will get:
                #              3,5,7,...(2n+1)
                #
                # Which is 4th, 6th,...(2n+2) columns
                # from Excel point of view. As it 
                # starts numbering from 1.
                #
                # So in Excel worksheet, even columns 
                # will have 1st test data.
                #####################################
                col = 1 + (sf*2)   
                row1 = row + 2
                
                ########################################
                # this function will write given list 
                # of numbers, row-wise 
                #
                #           Writing 2nd column
                ########################################
                
                # just titles: Time stamp & Test Description 
                ssef.write_excel_data( worksheet
                                      , col_2nd, row
                                      , col, 0, regular_10)
                                      
                # writing data exracted from Baseline file(s)
                ssef.write_excel_data_float( worksheet
                                            , final_iops_1st, row1
                                            , col, 0, regular_10)
                

                
                ##################################
                # 2n + 2 series, It will generate
                # even numbers.
                #
                # In our case, as sf starts from 1,
                # we will get:
                #              4,6,8,...(2n+2)
                #
                # Which is 5th, 7th,...(2n+3) columns
                # from Excel point of view. As it 
                # starts numbering from 1.
                #
                # So in Excel worksheet, odd columns 
                # will have 2nd test data from Same 
                # file.
                #####################################
                col = 2 + (sf*2)
                
                # Writing 3rd column
                # just titles: Time stamp & Test Description 
                ssef.write_excel_data( worksheet
                                      , col_3rd, row
                                      , col, 0, regular_10)
                
                # writing data exracted from Baseline file(s)
                ssef.write_excel_data_float( worksheet
                                      , final_iops_2nd, row1
                                      , col, 0, regular_10)
                                      
                
                ####################################
                # Calculating 
                #
                # Degradation is calculated with
                # respect to 1st file's data.
                ####################################
                degradation_1st_test = ssef.find_degradation(
                                                             avg_1st,
                                                             final_iops_1st
                                                             )
                
                degradation_2nd_test = ssef.find_degradation(
                                                             avg_2nd,
                                                             final_iops_2nd
                                                             )
                
                                  

                ##########################################
                # For statistical calculations
                #
                # It will calculate:
                #                  1.) Maximum
                #                  2.) Minimum
                #                  3.) Average
                #                  4.) Standard Deviation
                ##########################################
                
             
                #######################################
                # Writing Degradation and Staistics
                #
                #               1st Test 
                #######################################
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                            , degradation_1st_test, row2
                                            , col2, 0, regular_10)
                
                # conditional formatting
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_1st_test, test_name_list[1])                  
                
                
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)
                            
                                            
                ########################################
                # adding Series for 1st test
                # to the chart in "All Drives Baseline"
                ########################################
                chart.add_series({
                
                     #series name   
                    'name': str(test_name_list[1]), 
                    
                    # X-axis name
                    'categories': '='+str(ws_1st_name)+'!$A$' +str(row2+2)
                                    + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis values
                    'values': '='+str(ws_1st_name)+'!$' +str(svf.rank(col2+1))+
                                '$' +str(row2+2)+':$'+str(svf.rank(col2+1))+
                                '$'+str(row2+disk_no[0]+1),
                    
                    # Marker type            
                    'marker': {
                                'type': 'square',
                              }         
                    
                                    })                          
                    

                #######################################
                # Writing Degradation and Staistics
                #
                #               2nd Test 
                #######################################
                col2 +=1
                ssef.write_excel_data_float( worksheet
                                            , degradation_2nd_test, row2
                                            , col2, 0, regular_10)
                
                # conditional formatting
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_2nd_test, test_name_list[2])                  
                
                    
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)

                
                # adding Series for 2nd test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                
                    #series name
                    'name': str(test_name_list[2]),
                    
                    # X-axis name
                    'categories': '='+ str(ws_1st_name)+'!$A$' 
                                     + str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis name    
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))+
                                 '$'+str(row2+disk_no[0]+1),
                    # Marker type             
                    'marker': {
                                'type': 'circle',
                              }   
                              
                                    })
                                
                
                ####################################
                # Writing IOMeter Error data 
                ####################################

                # Writing 2nd column
                col3 += 1  
                ssef.write_excel_data( worksheet
                                      , final_errors_1st, row3
                                      , col3, 0, regular_10)
                # Writing 3rd column
                col3 +=1                      
                ssef.write_excel_data_float( worksheet
                                          , final_errors_2nd, row3
                                          , col3, 0, regular_10)
                
                ############################
                #       FOR LOOP END!      #
                #         BASELINE         #
                ############################
                             
                             
                             
        #########################################
        # Finding Average of all Baseline
        # and writing in Worksheet
        #
        #           1st Test 
        #########################################

        # Heading for Average, 1st Test
        avg_1st_text = ['Overall Baseline: \n'+str(just_test_name[0]), 'Calculated Average']

        # Writing IOps Error, 1st Test
        col = 1 + (sf_len*2) + 1
        row=0
        row+=4

        ssef.write_excel_data( worksheet
                              , avg_1st_text, row
                              , col, 0, regular_10)
         
        row1 = row + 2        
        ssef.write_excel_data_float( worksheet
                                    , avg_1st, row1
                                    , col, 0, regular_10)

        #########################################
        # Finding Average of all Baseline
        # and writing in Worksheet
        #
        #           2nd Test 
        #########################################
        
        # Heading for Average, 2nd Test
        avg_2nd_text = ['Overall Baseline: \n'+str(just_test_name[1]), 'Calculated Average']
        
        # Writing IOps Error, 1st Test
        col = 2 + (sf_len*2) + 1                            
        ssef.write_excel_data( worksheet
                              , avg_2nd_text, row
                              , col, 0, regular_10)
                 
        ssef.write_excel_data_float( worksheet
                                    , avg_2nd, row1
                                    , col, 0, regular_10)

        
        
        # Insert Chart in B2 Cell                          
        worksheet1.insert_chart('B2', chart)
        
        return [workbook, avg_1st, avg_1st_text, avg_2nd, avg_2nd_text,
        merge_format, bold_14, bold_12, bold_10, bold_10_l, regular_10, regular_10_l]
        
    
                #############################
                #                           #
#################    NON-OP Baseline END    #########################
                #                           #
                #############################

        
        
        
        
    ###############################
    #      NON - OP SUMMARY       #
    ###############################
    
    #################################################
    # 1.) This function will take Average of IOps
    # from Baseline files and calculate
    # degradation with respect to it.
    #
    # This function is similar to:
    # "generate_excel_performance". This function directly
    # extracts data from given file(s) list. While 
    # "generate_excel_performance" was extracting from
    # 2 files for "Office Vibe".
    #
    # 2.) It will also create and write
    # IOps, % Degradation & Error data
    # in "ws_1st_name" worksheet. 
    # Usually it is "Summary"    
    #
    # 3.) Now using this data. This function will
    # plot two charts. 
    #               a.) All Drives Vibe(ws_2nd_name)
    #               b.) High/Low/Average(ws_5th_name)
    ###################################################
        
    def generate_excel_non_op_baseline_summary(merge_format, bold_14, bold_12
                       , bold_10, bold_10_l, regular_10, regular_10_l, c_path
                       , sorted_files, file_dir_perf, workbook
                       , chassis_name, ws_1st_name, ws_2nd_name
                       , ws_5th_name, selector, avg_1st, avg_1st_text
                       , avg_2nd, avg_2nd_text, title):
    
        import os
        import xlsxwriter
        import sv_functions
        
        svf = sv_functions.SV_Functions

        
        ###################################
        #  Importing from other Directory
        ###################################
        import sys
        import os
        import numpy as np

        
        ##################################
        # Inserting "SSE" path to import 
        # and use some of its functions
        ##################################        
        os.chdir('..')
        c_path = os.getcwd()
        sys.path.insert(0, r''+str(c_path)+'/SSE')

        import sse_functions
        ssef= sse_functions.SSE_Functions
        
        
        ##################################
        # Jumping back to "SV" directory
        # This is our original directory
        ##################################
        sys.path.insert(0, r''+str(c_path)+'/SV')                               
                
                

        # length/number of sorted baseline files
        sf_len = len(sorted_files) 

        
        #############################
        # Create Baseline Worksheet
        #############################
        worksheet = workbook.add_worksheet(''+str(ws_1st_name))
        
        # Set column width to 30 points
        worksheet.set_column('A:Z', 30)

        
        # Setting borders
        regular_10.set_border(1)
        bold_10.set_border(1)
        regular_10_l.set_border(1)
        
        format1 = workbook.add_format({'bold':True, 'align': 'center'})
        format1.set_num_format('0.0')
        format1.set_border()

        
        #########################################
        # Merge 1st & 2nd rows, over columns.
        #########################################
        worksheet.merge_range('A1:Z1', 'VIBRATION TEST RESULTS', merge_format)
        worksheet.merge_range('A2:Z2', ''+str(title), merge_format)
        
        
        #########################################
        # Including Title for IOps Result,
        # "col" & "row" indicates position 
        # in the Excel sheet.
        #########################################
        col = 0
        row = 3
        worksheet.write(row, col, 'IOps Results', bold_12)
        
        
        ################################
        # Writing "All Drives Baseline"                          
        ################################                                              
        worksheet1 = workbook.add_worksheet(''+str(ws_2nd_name))                          
        

        #########################################
        # Chart Settings
        #########################################        
        # chart type
        chart = workbook.add_chart({'type': 'line'})
        
        # chart scale
        chart.set_size({'x_scale': 3, 'y_scale': 2.5}) 
        
        # chart title
        chart.set_title({
                         'name':''+str(chassis_name)
                         +' Basline Testing'\
                           '\nwith '+str(title)+
                            '\nIOMeter Test Results by Drive '
                        })
        
        # chart y axis range and Axis name                        
        chart.set_y_axis({
                            'name': 'IOps % Degradation',
                            'min': -15,
                            'max': 15
                         })
                         
        # chart x axis range, label position, gridlines, and Axis name                        
        chart.set_x_axis({
                            'label_position': 'low',
                            'position_axis': 'on_tick',
                            'min': 1,
                            'major_gridlines': 
                                            {
                                            'visible': True,
                                            'line': {'width': 0.5}
                                            }
                         })

        format1 = workbook.add_format({'bold':True, 'align': 'center'})
        format1.set_num_format('0.0')
        format1.set_border()
                 

        ###################################
        # This loop will read and extract  
        # Baseline file one by one. 
        #
        # It will write IOps data, 
        # Degradation, IOMeter Errors for 
        # both tests.
        ###################################

        for sf in range(sf_len):
            
            ####################################################
            # This function will extract all data 
            # from couple of Office Vibe .csv provided file(s).
            ####################################################
            # This function will extract all data 
            # from Single .csv Baselie file.
            ####################################################
            [final_iops_1st, final_errors_1st
            , final_iops_2nd, final_errors_2nd
            , final_drives, test_name_list
            , final_time_stamps, disk_no
            , just_test_name] = svf.extract_baseline_data(file_dir_perf,
                                                          sorted_files[sf],
                                                          selector)

                                                          
            ######################################                                                          
            # Joining/Modifying List so that we 
            # can directly write it in 1st Column
            ######################################
            
            # 1st column,IOps
            col_1st = [final_time_stamps[0]] + [test_name_list[0]] + final_drives
            
            # 2nd column,IOps
            col_2nd = [final_time_stamps[1]] + [test_name_list[1]] 
            
            # 3rd column,IOps
            col_3rd = [final_time_stamps[2]] + [test_name_list[2]]
                
        
        
            ##############################
            # We need to write 1st column
            # just once. So if it is a 
            # 1st file. Write 1st, 2nd, 
            # and 3rd column.
            ##############################
            if sf == 0:
            
                # store data of 1st file to find degradation value later
                file1_file_1st = final_iops_1st 
                file1_file_2nd = final_iops_2nd
                
                
                ######################################################
                # creating an empty list of zeros, which will finally
                # hold row-wise summation of All .csv files.
                # Which is used to calculate the Average
                ######################################################
                
                # 1st Test
                null_array_1st = np.zeros(int(disk_no[0])+1)
                null_array_1st = null_array_1st.tolist() 
                
                # 2nd Test
                null_array_2nd = null_array_1st
                
                
                #########################################
                # Zeros array to calculate average
                #########################################
                # Summing up 1st test data
                null_array_1st = np.sum([final_iops_1st, null_array_1st], axis = 0)

                # Summing up 2nd test data
                null_array_2nd = np.sum([final_iops_2nd, null_array_2nd], axis = 0)
                      
                
                                
                #############################################                    
                # Writing 1st column, IOps, Left justified     
                # Originally row = 3, for IOps title.
                # Now row =4, to write it from next 
                # row, col = 0 (1st column)
                #############################################
                
                #####################
                # Drives
                #####################
                row+= 1 
                ssef.write_excel_data( worksheet
                                      , col_1st, row
                                      , col, 0, regular_10_l)


                
                
                #####################
                # Baseline Average
                #####################
                
                # Average of 1st test from Baseline 
                col+=1
                row1 = row + 2
                
                # used for Plotting in High/Low/Average 
                td_row = row + 1 
                
                # Title of 1st Test, from Baseline data
                ssef.write_excel_data( worksheet
                                      , avg_1st_text, row
                                      , col, 0, regular_10)
                
                # Average IOps of 1st Test, from Baseline data
                ssef.write_excel_data_float( worksheet
                                            , avg_1st, row1
                                            , col, 0, regular_10)
                        
                               
                # Title of 2nd Test, from Baseline data 
                col+=1
                ssef.write_excel_data( worksheet
                                      , avg_2nd_text, row
                                      , col, 0, regular_10)

                # Average IOps of 2nd Test, from Baseline data
                ssef.write_excel_data_float( worksheet
                                            , avg_2nd, row1
                                            , col, 0, regular_10)
                        
                
                                
                                
                #####################
                #  PERFORMANCE DATA
                #####################

                # Writing 2nd column, center justified
                col+=1  # Moving to next Column
               
               
                # Test 1, IOps title (Time stamp & Test Description )
                # this function will write given list 
                # of strings row-wise
                ssef.write_excel_data( worksheet
                                      , col_2nd, row
                                      , col, 0, regular_10)
 
                
                # Test 1, IOps data
                # this function will write given list 
                # of numbers, row-wise                      
                ssef.write_excel_data_float( worksheet
                                          , final_iops_1st, row1
                                          , col, 0, regular_10)

                       
                       
                # Writing 3rd column
                # Test 2, IOps title (Time stamp & Test Description )
                col+=1                      
                ssef.write_excel_data( worksheet
                                      , col_3rd, row
                                      , col, 0, regular_10)
                # Test 2, IOps data
                ssef.write_excel_data_float( worksheet
                                            , final_iops_2nd, row1
                                            , col, 0, regular_10)
                
            
                ####################################
                # Writing degradation data 
                ####################################
                # calculate next row to write data
                #####################################
                
                ###################################################
                #
                # starting row (row2)=   row1(staring IOps row) 
                #                       + (number of disks + system) 
                #                       + 4 Blank Rows  
                #
                ####################################################
                row2 = row1 + (disk_no[0]+1) + 4 # 4 Blank Rows
                col2 = 0 
                
                # Degradation title
                worksheet.write(row2, col2, 'IOps, % degradation, compared to overall baseline tests', bold_12)
                
                # Writing 1st column, Degradation 
                # Drives
                row2 += 1
                ssef.write_excel_data( worksheet
                                      , col_1st[2:], row2
                                      , col2, 0, regular_10_l)
                
                
                #####################################
                # As degradation is calculated with 
                # respect to 1st file's values
                # leave 2nd and 3rd column blank.
                #
                # Write Statistical titles in 3rd column
                #########################################    
                col2 += 2
                row2_new = row2 + (disk_no[0]+1)
                
                start_col = col2
                # To write Statistical data
                stat_list = ['High', 'Low', 'Average', '1 σ']
                
                ssef.write_excel_data( worksheet
                                      , stat_list, row2_new
                                      , col2, 0, bold_10_l)
                
                
                ####################################
                # Degradation is calculated with
                # respect to 1st file's data.
                ####################################
                
                ################################
                # Comparing test with baseline 
                # so that Degradation is calculated
                # with correct Baseline average 
                # Test.
                ################################
                import re
                
                # exracting just numbers to compare
                jtn_1st = re.findall(r'\d+', just_test_name[0])
                jtn_2nd = re.findall(r'\d+', just_test_name[1])
                
                
                at_1st = re.findall(r'\d+', avg_1st_text[0])
                at_2nd = re.findall(r'\d+', avg_2nd_text[0])
                
                
                # 1st Test
                # if name of 1st same as 1st test's name from Baseline
                if int(jtn_1st[0]) == int(at_1st[0]):  
                    degradation_1st_test = ssef.find_degradation(avg_1st, final_iops_1st)
                
                # if name of 1st same as 2nd test's average name from Baseline
                elif int(jtn_1st[0]) == int(at_2nd[0]):
                    degradation_1st_test = ssef.find_degradation(avg_2nd, final_iops_1st)
                
                # else notify User about unmatch
                else:
                    print('\nTest name of Baseline and Performance file does not match!'\
                          '\nPlease make sure it matches, to generate a report.')

                      
                # 2nd test          
                # if name of 2nd test same as 1st test's name from Baseline
                if int(jtn_2nd[0]) == int(at_1st[0]):
                    degradation_2nd_test = ssef.find_degradation(avg_1st, final_iops_2nd)
                
                # if name of 2nd test same as 2nd test's average name from Baseline
                elif int(jtn_2nd[0]) == int(at_2nd[0]):
                    degradation_2nd_test = ssef.find_degradation(avg_2nd, final_iops_2nd)
                
                # else notify User about unmatch
                else:
                    print('\nTest name of Baseline and Performance file does not match!'\
                          '\nPlease make sure it matches, to generate a report.')
                          
                
                # Write degradation of 1st test
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                      , degradation_1st_test, row2
                                      , col2, 0, regular_10)
                
                # conditional formatting
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_1st_test, test_name_list[1])                  
                
                                      
                ##########################################
                # For statistical calculations
                #
                # It will calculate:
                #                  1.) Maximum
                #                  2.) Minimum
                #                  3.) Average
                #                  4.) Standard Deviation
                ##########################################
                
                # For 1st test
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)
                
                # adding Series for 1st test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                
                    #series name
                    'name': str(test_name_list[1]),
                    
                    # X-axis name
                    'categories': '='+str(ws_1st_name)+'!$A$' 
                                     +str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    # Y-axis values                        
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))
                                 +'$'+str(row2+disk_no[0]+1),
                    
                    # Marker type
                    'marker': {
                                'type': 'square',
                              }         
                                    })                          
                
                
                #########################
                # 2nd Test 
                #########################
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                      , degradation_2nd_test, row2
                                      , col2, 0, regular_10)
                
                # conditional formatting
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_2nd_test, test_name_list[2])                  
                               
                # For 2nd test, Staistics                  

                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)
                
                # adding Series for 2nd test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                    
                    #series name
                    'name': str(test_name_list[2]),
                    
                    # X-axis name
                    'categories': '='+ str(ws_1st_name)+'!$A$' 
                                     + str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                     
                    # Y-axis values
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))
                                 +'$'+str(row2+disk_no[0]+1),
                    
                    # Marker type
                    'marker': {
                                'type': 'circle',
                              }   
                              
                                })
                
                
                ####################################
                # Writing IOMeter Error data 
                ####################################
                # calculate next row to write data
                #####################################
                
                ###################################################
                #
                # starting row (row3)=   row2(staring Degradation row) 
                #                       + (number of disks + system) 
                #                       + 6 Blank Rows  
                #
                ####################################################
                row3 = row2 + (disk_no[0]+1) + 6 
               
                
                # Error Title
                col3 = 0 
                worksheet.write(row3, col3, 'IOMeter Errors', bold_12)
                
                # Writing 1st column, Degradation, Drives #     
                row3 += 1
                ssef.write_excel_data( worksheet
                                      , col_1st[2:], row3
                                      , col3, 0, regular_10_l)
                
                
                # Writing 2nd column, 1st Test Errors
                col3 = 2
                col3 += 1  
                      
                ssef.write_excel_data( worksheet
                                      , final_errors_1st, row3
                                      , col3, 0, regular_10)
                
                # Writing 3rd column, 2nd Test Errors
     
                col3+=1                      
                ssef.write_excel_data_float( worksheet
                                          , final_errors_2nd, row3
                                          , col3, 0, regular_10)

            
            
            
            ##############################
            #     PERFORMANCE DATA
            ##############################
            # We need to write 1st column
            # just once. So if it is a 
            # 1st file. Write 1st, 2nd, 
            # and 3rd column.
            #
            # Else just write 2nd and 3rd 
            # column from 2nd and rest of 
            # files.
            ##############################
                                         
            else:
            
                #########################################
                # Null array to calculate average
                #########################################
                
                # Summing up 1st test data
                null_array_1st = np.sum([final_iops_1st, null_array_1st], axis = 0)

                # Summing up 2nd test data
                null_array_2nd = np.sum([final_iops_2nd, null_array_2nd], axis = 0)
         
         
                # this function will write given list 
                # of strings row-wise
                row1 = row + 2 
                
                
                ##################################
                # 2n + 3 series, It will generate
                # odd numbers.
                #
                # In our case, as sf starts from 1,
                # we will get:
                #              5,7,9...(2n+1)
                #
                # Which is 6th,8th...(2n+3) columns
                # from Excel point of view. As it 
                # starts numbering from 1.
                #
                # So in Excel worksheet, even columns 
                # will have 1st test data.
                #####################################
                col = 3 + (sf*2)         
                
                # 1st Test, title (Time stamp & Test Description)
                ssef.write_excel_data( worksheet
                                      , col_2nd, row
                                      , col, 0, regular_10)
                
                
                # this function will write given list 
                # of numbers, row-wise 
                # Writing 2nd column
                
                # 1st Test, IOps data
                ssef.write_excel_data_float( worksheet
                                              , final_iops_1st, row1
                                              , col, 0, regular_10)
                         
                
                ##################################
                # 2n + 4 series, It will generate
                # even numbers.
                #
                # In our case, as sf starts from 1,
                # we will get:
                #              6,8,10,...,(2n+5)
                #
                # Which is 7th, 9th, ...,(2n+5) columns
                # from Excel point of view. As it 
                # starts numbering from 1.
                #
                # So in Excel worksheet, odd columns 
                # will have 2nd test data from Same 
                # file.
                #####################################
                col = 4 + (sf*2)                       
                
                # Writing 3rd column, 
                # 2nd Test, title (Time stamp & Test Description)
                ssef.write_excel_data( worksheet
                                      , col_3rd, row
                                      , col, 0, regular_10)
                      
                # 2nd Test, IOps data
                ssef.write_excel_data_float( worksheet
                                      , final_iops_2nd, row1
                                      , col, 0, regular_10)
                                      
                
                ####################################
                # Calculating 
                #
                # Degradation is calculated with
                # respect to Average of that test.
                ####################################
                
                ################################
                # Comparing test with baseline 
                # so that Degradation is calculated
                # with correct Baseline average 
                # Test.
                ################################
                
                jtn_1st = re.findall(r'\d+', just_test_name[0])
                jtn_2nd = re.findall(r'\d+', just_test_name[1])
                
                
                at_1st = re.findall(r'\d+', avg_1st_text[0])
                at_2nd = re.findall(r'\d+', avg_2nd_text[0])
               
                
                # 1st Test
                # if name of 1st same as 1st test's name from Baseline
                if int(jtn_1st[0]) == int(at_1st[0]):
                    degradation_1st_test = ssef.find_degradation(avg_1st, final_iops_1st)
                
                # if name of 1st same as 2nd test's average name from Baseline
                elif int(jtn_1st[0]) == int(at_2nd[0]):
                    degradation_1st_test = ssef.find_degradation(avg_2nd, final_iops_1st)
                
                # else notify User about unmatch
                else:
                    print('\nTest name of Baseline and Performance file does not match!'\
                          '\nPlease make sure it matches, to generate a report.')
                      


                      
                # 2nd test   
                # if name of 2nd test same as 1st test's name from Baseline                
                if int(jtn_2nd[0]) == int(at_1st[0]):
                    degradation_2nd_test = ssef.find_degradation(avg_1st, final_iops_2nd)
                
                # if name of 2nd test same as 2nd test's average name from Baseline
                elif int(jtn_2nd[0]) == int(at_2nd[0]):
                    degradation_2nd_test = ssef.find_degradation(avg_2nd, final_iops_2nd)
                
                 # else notify User about unmatch
                else:
                    print('\nTest name of Baseline and Performance file does not match!'\
                          '\nPlease make sure it matches, to generate a report.')
                
                                  

                ##########################################
                # For statistical calculations
                #
                # It will calculate:
                #                  1.) Maximum
                #                  2.) Minimum
                #                  3.) Average
                #                  4.) Standard Deviation
                ##########################################
               
                #######################################
                # Writing Degradation and Staistics
                #
                #               1st Test 
                #######################################
                col2 += 1
                ssef.write_excel_data_float( worksheet
                                      , degradation_1st_test, row2
                                      , col2, 0, regular_10)
                
                # conditional formatting
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_1st_test, test_name_list[1])                  
                
                
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)
                
                
                # adding Series for 1st test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                    
                    #series name
                    'name': str(test_name_list[1]),
                    
                    # X-axis
                    'categories': '='+str(ws_1st_name)+'!$A$' 
                                     +str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis    
                    'values': '='+str(ws_1st_name)+'!$'+str(svf.rank(col2+1))
                                 +'$' +str(row2+2)+':$'+str(svf.rank(col2+1))
                                 +'$'+str(row2+disk_no[0]+1),
                    
                    # marker type
                    'marker': {
                                'type': 'square',
                              }         
                                 
                                 })                          
                    
                
                
                
                #######################################
                # Writing Degradation and Staistics
                #
                #               2nd Test 
                #######################################
                col2 +=1
                
                # degradation data
                ssef.write_excel_data_float( worksheet
                                      , degradation_2nd_test, row2
                                      , col2, 0, regular_10)
                
                # conditional formatting
                [workbook, worksheet] = svf.perform_CF_on_list(ssef, workbook, worksheet, row2, col2, degradation_2nd_test, test_name_list[2])                  
                
                # Calculating Maximum, Minimum, Average, and Standard Deviation
                # from Degradation list
                svf.calculate_n_write_stats(worksheet, row2_new, row2, col2, format1)

                
                # adding Series for 2nd test
                # to the chart in "All Drives Baseline"
                chart.add_series({
                    
                    #series name
                    'name': str(test_name_list[2]),
                        
                    # X-axis
                    'categories': '='+str(ws_1st_name)+'!$A$' 
                                     +str(row2+2)
                                     + ':$A$'+str(row2+disk_no[0]+1),
                    
                    # Y-axis    
                    'values': '='+str(ws_1st_name)+'!$'
                                 +str(svf.rank(col2+1))+'$' +str(row2+2)
                                 +':$'+str(svf.rank(col2+1))+'$'+str(row2+disk_no[0]+1),
                    
                    # marker type
                    'marker': {
                                'type': 'circle',
                              }   
                              
                                })
                                
                
                ####################################
                # Writing IOMeter Error data 
                ####################################

                # Writing 2nd column
                col3 += 1  
                ssef.write_excel_data( worksheet
                                      , final_errors_1st, row3
                                      , col3, 0, regular_10)
                # Writing 3rd column
                col3 +=1                      
                ssef.write_excel_data_float( worksheet
                                          , final_errors_2nd, row3
                                          , col3, 0, regular_10)
                
                ############################
                #       FOR LOOP END!      #
                ############################
        
                     
        # Inserting "Line chart" after 
        # Loop ends at B2
        worksheet1.insert_chart('B2', chart)
        
 

        #########################
        # Creating 
        # Hi, Lo, Avg Chart in 
        # a given Workbook.
        #########################
        worksheet = workbook.add_worksheet(''+str(ws_5th_name))

        chart = workbook.add_chart({'type': 'stock'})
        
        chart.set_size({'x_scale': 3 , 'y_scale': 2.5}) 
        
        # chart title
        chart.set_title({'name':''+str(chassis_name)
                                  +' Testing'\
                                   '\nwith '+str(title)+
                                   '\nIOMeter Test Results, High/Low/Average'})
        
        # chart y axis range and Axis name                        
        chart.set_y_axis({'name': 'IOps % Degradation',
                         'min': -20, 'max': 20})
                         
        # chart x axis range, label position, gridlines, and Axis name                        
        chart.set_x_axis({
                            'label_position': 'low'
                          , 'position_axis': 'on_tick'
                          , 'min': 1
                          , 'major_gridlines': {
                                                'visible': True,
                                                'line': {'width': 0.5}
                                                }
                         })

                         
        ####################
        # Adding Series
        ####################
        # High
        ###################
        chart.add_series({
                            # shape
                            'line': 'o',
                            
                            # name
                            'name': 'High',
                            
                            # X-axis
                            'categories': '='+str(ws_1st_name)
                                             +'!$'+str(svf.rank(start_col+2)) 
                                             +str(td_row+1)
                                             + ':$'+str(svf.rank(col2+1))
                                             +'$'+str(td_row+1),
                            # Y-axis    
                            'values': '='+str(ws_1st_name)
                                         +'!$'+str(svf.rank(start_col+2)) 
                                         +str(row2_new+1)
                                         + ':$'+str(svf.rank(col2+1))
                                         +'$'+str(row2_new+1),
                     
                            'line':   {'none': True},    
                            #'line':       {'color': 'red'},
                            'marker': {
                                        'type': 'diamond',
                                        'border': {'color': 'red'},
                                        'fill':   {'color': 'red'},
                                    }

                        })
                        
                        
        #################
        # Low
        #################
        chart.add_series({
                            # series name
                            'name': 'Low',
                            
                            # X-axis
                            'categories': '='+str(ws_1st_name)
                                             +'!$'+str(svf.rank(start_col+2)) 
                                             +str(td_row+1)
                                             + ':$'+str(svf.rank(col2+1))
                                             +'$'+str(td_row+1),
                            # Y-axis    
                            'values':     '='+str(ws_1st_name)+'!$'+str(svf.rank(start_col+2)) 
                                            +str(row2_new+2)
                                            + ':$'+str(svf.rank(col2+1))+'$'+str(row2_new+2),
                     
                                
                            'line':   {'none': True},
                           
                            # Marker properties
                            'marker': {
                                        'type': 'square',
                                        'border': {'color': 'blue'},
                                        'fill': {'color':'blue'}
                                      }
                          })

        #################
        # Average
        #################
        chart.add_series({
                            # series name
                            'name': 'Average',
                            
                            # X-axis
                            'categories': '='+str(ws_1st_name)+'!$'
                                             +str(svf.rank(start_col+2)) 
                                             +str(td_row+1)
                                             + ':$'+str(svf.rank(col2+1))
                                             +'$'+str(td_row+1),
                            # Y-axis    
                            'values': '='+str(ws_1st_name)+'!$'
                                         +str(svf.rank(start_col+2)) 
                                         +str(row2_new+3)
                                         + ':$'+str(svf.rank(col2+1))+
                                         '$'+str(row2_new+3),
                            # line property
                            'line':   {'none': True},
                            
                            # marker properties
                            'marker': {
                                        'type': 'square',
                                        'border': {'color': 'green'},
                                        'fill': {'color':'green'}
                                      }
                        })
                        
                        
        
        
        # Insert High/Low/Average chart at B2
        worksheet.insert_chart('B2', chart)
        
        return workbook
        
                    #############################
                    #                           #
#####################       Op SUMMARY END      #########################
                    #                           #
                    #############################

                    
                    
#############################
#           END             #
#############################                      