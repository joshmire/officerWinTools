# officerWinTools R package  
THIS PACKAGE REQUIRES A WINDOWS OPERATING SYSTEM AND MICROSOFT OFFICE.  
  
The functions in the package run VBScript in Command Prompt.  
  
The package includes functions for exporting Microsoft Word documents (`.docx`) and PowerPoint 
presentations (`.pptx`) as Portable Document Format (`.pdf`) files and a function for updating the 
tables of contents of a Microsoft Word document.  
  
 > *Functions to compliment [officer](https://github.com/davidgohel/officer) R package using 
 Microsoft Office in a Windows environment*  
  
 ## Installation  
 The package is currently only available on GitHub:  
 ```r
 devtools::install_github("joshmire/officerWinTools")
 ```
  
 ## Development Plan
 *Version 1.0.0*:  Add `insert_slide()` function which will copy a range of slides from one 
 PowerPoint presentation into another at a specified index or else at the end of the presentation
 
 *Version 2.0.0*:  Remove need for writing `temp.vbs`.
 
 *Version 3.0.0*:  Linux compatibility.
 
 *Archive*:  If Linux compatibility can be achieved, this package will be archieved and re-uploaded
 as officerTools and submitted to CRAN.
  
 ## Author's note  
The package is meant to be used along side the [officer](https://github.com/davidgohel/officer) R 
package by [David Gohel et al](https://davidgohel.github.io/officer/authors.html).  The author of 
the package has no affiliation with the [authors](https://davidgohel.github.io/officer/authors.html) 
of [officer](https://github.com/davidgohel/officer).  
