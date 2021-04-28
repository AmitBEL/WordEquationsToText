Script Purpose
==============  
Given a docx file containing equations  
Generate new docx file where all the equations are converted to text  
The new docx file is named as the original docx file name + '_f'  

Script Logic  
--------------  
  1. Get the docx file path as parameter
  2. Copy the docx file and convert the duplicated file to zip file
  3. Extract the zip file
  4. Change all the equations to text in document.xml
  5. Archive the new files to zip file
  6. Convert the zip file to docx file
