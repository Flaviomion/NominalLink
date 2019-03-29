# NominalLink
A script to link value between excel files based on Nominal Parameter.
A VBscript that open an excel file, that will be the Destination of all value, and search for link created on Comment of the Value Cell.
The link must have this format:
  Flink#D#dir=X:\yyy#file=Analisis#ParameterName#Elink
  Flink# and #Elink     characterizes the link
  #D#                   stands for destination
  dir=X:\yyy            the directory where the script search the files that are sources of the values
  file=Analisis         Opzional the file name of file that is sources of the values
  ParameterName         the Name of the parameterthat will be copied from source to destination.
  
Then the script look all links and try to find sources from the file and from directory find on the link.
The match is done by check that Destination ParameterName and Source ParameterName are the same reading link on source files with     this format:
  Flink#S#ParameterName#Elink
  Flink# and #Elink     characterizes the link
  #S#                   stands for sources
  ParameterName         the Name of the parameter that will be copied to destination.
  
If you have on directory c:\testNominalLink a file named Source*.xlsx and on c:\report a file named report.xlsx
to prepare the excel file for te migration of value you need to put in file c:\testNominalLink on the comment of the source cell the line:
  Flink#S#pippo#Elink
  
and in c:\report\report.xlsx put next line on the comment of the destination cell:
  Flink#D#dir=C:\testNominalLink#file=Source#pippo#Elink
  
 and start the NominalLink.vbs, ath the end of the script a report.html is displayed with the information of what was done.
 
  
