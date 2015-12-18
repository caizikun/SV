####################################################
#                                                  #
#   Instructions to run SV scripts in Python       #
#                                                  #
#   Author: Zankar Sanghavi                        #
#                                                  #
#   © Dot Hill Systems Corporation                 #
#                                                  #
####################################################



####################################################
#                                                  #
#            Installation and running              #
#						   #
#  If you already installed Python 3 and Packages  #
#       then you can start from step:3  	   #
#                                                  #
####################################################

1.) Before running, you need to Install Anaconda Python 3
    from Continuum website or you can also use a Setup file 
    stored in "Setup Files" folder.

    Path: meng:\Hard Drive\Zankar Sanghavi\Setup Files\Anaconda3-2.2.0-Windows-x86_64

2.) Once you install Python, you need to install 3 Librabies
    or Packages (comtypes-1.1.1, openpyxl-2.2.5 and 
    XlsxWriter-0.7.3) to support these scripts. 
    
    To install a package:
    
    a.) Copy and Paste "Packages" in your local directory.
    
    b.) Open Command Prompt from Start Menu
    
    c.) Go to that directory by using this command:
    
        cd Absolute Path of Package
        
        For example:
        
            cd C:\Packages\comtypes-1.1.1
                
    d.) Then enter:
                
            ipython setup.py install
            
    Repeat process from Step (c) for each Package.
    
3.) Now you are ready to run a Report Automation script for SV 
    Test. Change the path to Specific path and enter. 
    
    ipython.

4.) Once you are in iPython, you can simply run script by entering 
    this command:
    
    run main_sv.py

5.) Then you will need to Enter full path of any . CSV SV files (i.e. Performance 
    file), Chassis name and name of Report.



####################################################
#                                                  #
#           		 NOTE  			   #
#                                                  #
####################################################

In order to generate a Report, User should take care of this before hand.

	i.) User should enter a file with extension ".csv"
	
	ii.) For OP SV Report generation, baseline file name should have these qualifiers:

		a.) Pre - In the beginning.
		b.) Baseline - In the end.

	   It will extract the name of Test using these strings. So whatever is situated between
	   first number and "Baseline" will be kept in generated Report.

		Example: 1_Y_Pre_Office Vibe_Opt2_Baseline

		In report: Y_Pre_Office Vibe_Opt2

	iii.) For OP Sv Report generation. Final Baseline should have: "Final" and "SV" qualifiers. It can 
	     not have "Post" and "Drop" qualifiers.

	iv.) For OP SV Report generation. Performance files should not contain "Baseline" & "Post"
	     strings in its name.

	     Moreover, it should contain: "Perf" in the end or in-between to extract the file name.

		Example: 6_Y_Op_Shock_5g_Perf

		In report: Y_Op_Shock_5g 	

	v.) For OP SV report generation. Swept sine performance file should
	contain "SS" qualifier.
	
	Op shock performance files should have atleast have any one of these { "Op Shock"
	, "OP Shock", "op Shock", "OP shock", "op shock"} qualifier. 

	And Office vibe performance file should have "Office" or "office" qualifier.

	This method will make sure, that data is written in following sequence:
		1.) Swept sine	
		2.) Office Vibe (2 individual test files)
		3.) Op shock

	v.) If there are no performance file(s) in the directory, the process will give a Warning 
	    and Terminate.

	vi.)  For NON-OP SV Report generation. 

		BASELINE FILES FOR "BASELINE" & "ALL DRIVES BASELINE" WORKSHEET
		
		Baseline file # 1: Exclude: "Post" & "Drop"
		 		   Include: "1"
		
		Baseline file # 2: Exclude: "Post" & "Drop"
				   Include: "Final"& "SV"
		
		Baseline file # 3: Exclude: "Post" & "Drop"
				   Include: "Final"

		If any of these files are not present it won't appear in the report.

	vii.) For NON-OP SV Report generation.
	
	      To qualify files for "SUMMARY", "ALL DRIVES VIBE", and "HI/LOW/AVG CHART" WORKSHEET
		
	      It should have " 7, 9, 15, 17, 23, 25" qualifier in it. If this file is absent, it 
	      won't appear in the report. 



####################################################
#                                                  #
#            Error & possible Cause                #
#                                                  #
#################################################### 

1.) ImportError : It takes adavantages of some functions from "SSE" & "Common Scripts" directory.  
	          So if can not find those directories, then it will throw this error. If directory
		  name is changed then it needs to changed in Scripts as well.

2.) ValueError  : It is possible that number of drives of initial file does not match with number
		  of drives in later files. So it is impossible to compare different number of 
		  drives, so it throw this error.
	


####################################################
#                                                  #
#            Tips and Tricks for the User          #
#                                                  #
####################################################

1.) To change the directory, you can just Copy path from Address bar 
    and use right mouse click to paste it in a Command prompt window.

	Example: cd path; cd C:\SV_Python\SV

2.) After you get in to iPython, you can Browse files by entering 
    initial characters and then pressing Tab.
    
    Example:
    
    run  main_s -> Tab -> run main_sv.py 

	
3.) For files path you have to make sure that file exits, or a simple way
    get a path is to Drag and drop the file on Command Prompt window. You 
    can apply this where ever it says Enter path of a File.



####################################################
#                                                  #
#                   Know Issues                    #
#                                                  #
####################################################    

None, till 10/22/2015
