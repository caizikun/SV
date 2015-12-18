####################################################
#                  Revision: 1.0                   #
#              Updated on: 11/06/2015              #
####################################################

####################################################
#                                                  #
#   This script calls all functions                #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################
import sys
import os
import sv_functions
svf = sv_functions.SV_Functions

import extract_create_baseline
ecb = extract_create_baseline.Extract_Create_Baseline

c_path = os.getcwd()



###################################
#  To get path of current working 
#  directory.
###################################
path_1st = input('\nPlease enter full path of any CSV file: ')

#to check if it is a .csv file
while 'csv' not in path_1st:
    print('\nIt is not a .csv file!')
    path_1st = input('\nPlease enter full path of any CSV file: ')

    
# For title in Excel Charts
chassis_name = input('\nPlease enter Chassis Name: ')
    
    
# to save report name    
report_name = input('\nSave Report as: ')    


# check time to calculate processing time
import time
import os

start_time = time.time()
    
    
# get File's path
file_dir_perf= os.path.dirname(r''+str(path_1st)) #directory path
file_dir_perf= str(file_dir_perf).replace('"','') #remove "" quotes

file_list_perf=os.listdir(file_dir_perf) # find files in this directory



#################################################################
# Filter of 'baseline', 
# just to read baseline .csv 
# files.
#
# "filter" function takes:
#                          1.) List of file(s)
#                          2.) Qualifier string
#                          3.) 0 - include files with qualifier
#                              1 - exclude files with qualifier
#################################################################

##############################
#   Qualify OP SV files      #
##############################

# Include file(s) with "Baseline"
file_list_pre = svf.filter(file_list_perf, 'Baseline', 0)

# Include file(s) with "Pre" - Precondition files
file_list = svf.filter(file_list_pre, 'Pre', 0)

# Exclude File name containinig 'Non', 'NEBS', 'Post', & 'Drop'
file_list = svf.filter(file_list, 'Non', 1)
file_list = svf.filter(file_list, 'NEBS', 1)
file_list = svf.filter(file_list, 'Post', 1)
file_list = svf.filter(file_list, 'Drop', 1)
file_list = svf.filter(file_list, 'Perf', 1)


# Condition to qualify Final Baseline file
file_list_final = svf.filter(file_list_pre, 'Post', 1) # exclude

file_list_final = svf.filter(file_list_final, 'Perf', 1) # include

file_list_final = svf.filter(file_list_final, 'Drop', 1) # exclude

file_list_final = svf.filter(file_list_final, 'Final', 0) # include

file_list_final = svf.filter(file_list_final, 'SV', 0) # include

file_list_final = svf.filter(file_list_final, '.csv', 0) # include


# List of All Baseline files with Final Baseline File
final_files_list = file_list + file_list_final
#print(final_files_list)



############################
# Qualify Performance from
# all files list.
#
# "file_list1" is list of 
# final PERFORMANCE files
############################

# Exclude Baseline file(s)
file_list1 = svf.filter(file_list_perf, 'Baseline', 1) # exclude

file_list1 = svf.filter(file_list1, 'Post', 1) # exclude

file_list1 = svf.filter(file_list1, 'Drop', 1) # exclude

file_list1 = svf.filter(file_list1, '.csv', 0) # include

file_list1 = svf.filter(file_list1, 'Perf', 0) # include

# sort performance as per their number
sorted_perf = svf.sort_files(file_list1)


########################################
# If there are no performance file(s).
# It means insufficent data to generate
# report. 
#
# Another possibility is there is no 
# "Perf" string in Performance file(s).
# Notify User if this is the case and 
# terminate the process.
########################################
if not sorted_perf:
    print ('\n\nNo Performance file(s) found!'\
            '\n\nPlease make sure that Performance file(s)'\
            ' have "Perf" string in each file.'\
            '\n\n\nPROCESS TERMINATED !')
            
    sys.exit()
    
    
#########################################
# Performance file(s) should be in Y,X,Z 
# order. Qualify each axis and making a 
# list in required order.
#########################################            
x_axis_files = svf.filter(sorted_perf, 'X', 0)
y_axis_files = svf.filter(sorted_perf, 'Y', 0)
z_axis_files = svf.filter(sorted_perf, 'Z', 0)

perf_files_list = y_axis_files + x_axis_files + z_axis_files
#print(perf_files_list)


##############################
#   Qualify NON-OP SV files  #
##############################

# first file from the sorted file list
non_ops_baseline = svf.sort_files(file_list_pre)
b1 = [non_ops_baseline[0]]

# Condition to qualify Final Baseline file
file_list_final = svf.filter(non_ops_baseline, 'Post', 1) # exclude

file_list_final = svf.filter(file_list_final, 'Drop', 1) # exclude

file_list_final = svf.filter(file_list_final, 'Final', 0) # include

b2 = svf.filter(file_list_final, 'SV', 0) # include


# condition to qualify last Post drop file
file_list_final = svf.filter(non_ops_baseline, 'Post', 0) # include

file_list_final = svf.filter(file_list_final, 'Drop', 0) # include

b3 = svf.filter(file_list_final, 'Final', 0) # include


baseline_files = b1 + b2 + b3


##############################
# Qualify summary file
##############################

s1 = svf.filter(non_ops_baseline, '7', 0) # include
s2 = svf.filter(non_ops_baseline, '9', 0) # include
s3 = svf.filter(non_ops_baseline, '15', 0) # include
s4 = svf.filter(non_ops_baseline, '17', 0) # include
s5 = svf.filter(non_ops_baseline, '23', 0) # include
s6 = svf.filter(non_ops_baseline, '25', 0) # include

# Post Drop files

post_drop_files = svf.filter(non_ops_baseline, 'Post', 0) # include

post_drop_files = svf.filter(post_drop_files, 'Drop', 0) # include


# variable useful generate NON-OP report  
non_op = 1


try:
    final_summary_files = [s1[0]] + [s2[0]] + [s3[0]] + [s4[0]] + [s5[0]] + [s6[0]] +  post_drop_files
    #final_summary_files
except IndexError:
    print('\nInsufficient data for NON-OP Report!'\
           '\nOnly OP Report will be generated.')
    
    # set this variable to zero if not enough files are present  
    non_op = 0

#######################
#    EXCEL OP         #
#######################

###################################################
# Writing an empty Excel file, name given by User.
###################################################
import xlsxwriter
workbook = xlsxwriter.Workbook(r''+str(file_dir_perf)+
                               '\\' +str(report_name)+'_Op_SV_Qualification'
                               + '.xlsx')

                          


##########################
# Writing BASELINE data
##########################

#################################
# Sorting files so that we 
# extract data directly. And each 
# drives will be a sequence. 
#################################
sorted_files = svf.sort_files(final_files_list)
#print(sorted_files)

# Name of 1st Worksheet in a report
ws_1st_name = 'Baseline'

# Name of 2nd Worksheet in a report
ws_2nd_name = 'All Drives Baseline'

# Title for Charts and Worksheets for Ops
title = 'Rails and Blocks'

#######################################
# selector = 0 
# For Baseline file name extraction
#######################################
# This function will return Average of 
# both Tests to write it in "Summary"
# worksheet.
########################################
[merge_format
 , bold_14
 , bold_12
 , bold_10
 , bold_10_l
 , regular_10
 , regular_10_l
 , workbook
 , avg_1st
 , avg_1st_text
 , avg_2nd
 , avg_2nd_text] = ecb.generate_excel_baseline(c_path, sorted_files
                                               , file_dir_perf, workbook
                                               , chassis_name
                                               , ws_1st_name
                                               , ws_2nd_name
                                               , 0
                                               , title)
                          

                          
  
  
###########################
# Writing PERFORMANCE data 
###########################

ws_3rd_name = 'Summary'
ws_4th_name = 'All Drives Vibe'
ws_5th_name = 'Hi, Lo, Avg Chart'

workbook = ecb.generate_excel_performance(merge_format, bold_14, bold_12
                                         , bold_10, bold_10_l, regular_10
                                         , regular_10_l, c_path
                                         , perf_files_list
                                         , file_dir_perf, workbook
                                         , chassis_name, ws_3rd_name
                                         , ws_4th_name, ws_5th_name
                                         , 1 # selector
                                         , avg_1st, avg_1st_text
                                         , avg_2nd, avg_2nd_text
                                         , title)
                          

                                                    

                                                    

# close and save workbook                                                    
workbook.close() 


if non_op == 1:
    #######################
    #    EXCEL Non-Ops    #
    #######################
    ###################################################
    # Writing an empty Excel file, name given by User.
    ###################################################
    import xlsxwriter
    workbook1 = xlsxwriter.Workbook(r''+str(file_dir_perf)+
                                   '\\' +str(report_name)+'_Non-Op_SV_Qualification'
                                   + '.xlsx')

                              


    ##########################
    # Writing BASELINE data
    ##########################

    #################################
    # Sorting files so that we 
    # extract data directly. And each 
    # drives will be a sequence. 
    #################################

    # Title for Charts and Worksheets for Non-Op
    non_op_title = 'Direct Mount'

    #######################################
    # selector = 0 
    # For Baseline file name extraction
    #######################################
    # This function will return Average of 
    # both Tests to write it in "Summary"
    # worksheet.
    ########################################
    [workbook1,
     avg_1st,
     avg_1st_text,
     avg_2nd,
     avg_2nd_text,
     merge_format, 
     bold_14, 
     bold_12, 
     bold_10, 
     bold_10_l,
     regular_10,
     regular_10_l] = ecb.generate_excel_non_op_baseline(c_path, baseline_files
                                                    , file_dir_perf, workbook1
                                                    , chassis_name
                                                    , ws_1st_name
                                                    , ws_2nd_name
                                                    , 0
                                                    , non_op_title)



                                                   
    workbook1 = ecb.generate_excel_non_op_baseline_summary(merge_format, bold_14,
                                                           bold_12, bold_10, bold_10_l,
                                                           regular_10, regular_10_l,
                                                           c_path, final_summary_files,
                                                           file_dir_perf, workbook1, 
                                                           chassis_name, ws_3rd_name,
                                                           ws_4th_name, ws_5th_name,
                                                           0, avg_1st, avg_1st_text,
                                                           avg_2nd, avg_2nd_text,
                                                           non_op_title)
                                                   
                           
                  
    # close and save workbook                                                    
    workbook1.close()   
    
# go back to original directory
os.chdir(r''+str(c_path)) 



##########################                
# To show elapsed time   
##########################
elapse_time =round((time.time() - start_time),2) # seconds
if elapse_time < 60 :
    print("\nElapsed time: %s seconds" % elapse_time )
else:
    print("\nElapsed time: %s minutes" % round(((time.time() - start_time)/60),2))
    
    
    
#############################
#           END             #
#############################    