# NominalLink
A script to link value between excel files based on Nominal Parameter.
A VBscript that open an excel file, that will be the Destination of all value, and search for link created on Comment of the Value Cell.
The link must have this format:
  Flink#D#cartella=X:\yyy#file=Analisis#ParameterName#Elink
  
  Flink# and #Elink     characterizes the link
  #D#                   stands for destination
  cartella=X:\yyy            the directory where the script search the files that are sources of the values
  file=Analisis         Opzional the file name of file that is sources of the values
  ParameterName         the Name of the parameterthat will be copied from source to destination.
  
Then the script look all links and try to find sources from the file and from directory find on the link.
The match is done by check that Destination ParameterName and Source ParameterName are the same reading link on source files with     this format:
  Flink#S#ParameterName#Elink
  Flink# and #Elink     characterizes the link
  #S#                   stands for sources
  ParameterName         the Name of the parameter that will be copied to destination.
  
If you have on directory c:\testNominalLink a file named Source*.xlsx and on c:\report a file named report_01-01-2019.xlsx
to prepare the excel file for te migration of value you need 

1) Create a file NominalLink.ini where the first line is the directori were there is the Destination excel
   the second line is the first part of the name of the Destination Excel file (the remaining part of the name can be wath you want and canbe changed without problem)
   this is an examles of NominalLink.ini:
   
   c:\report
   report
   
2)  Put in file c:\testNominalLink\Source_01-01-2019.xlsx on the comment of the source cell the line:
  Flink#S#pippo#Elink
  
3) and in c:\report\report_01-01-2019.xlsx put next line on the comment of the destination cell:
  Flink#D#dir=C:\testNominalLink#file=Source#pippo#Elink
  
 and start the NominalLink.vbs, ath the end of the script a report.html is displayed with the information of what was done
 that is take the value on source cell (c:\testNominalLink\Source_01-01-2019.xlsx) and put it on destination cell (c:\report\report_01-01-2019.xlsx)
 
 Good linking.
 
  
