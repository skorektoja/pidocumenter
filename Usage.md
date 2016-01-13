# PI Diff #


To use the pi diff tool the following commandline should be called.
```
 Usage:
 php   <pathto>pidiff.php <directory> <old> <new> <format>

 <directory> Diretory where the files exists
 <old>       Filename of the old xim file
 <new>       Filename of the new xim file
 <format>    Output format 1 for EXCEL2007 , 2 for EXCEL2003
```

An example on this could be the following.
```
C:\Projects\test\diff2>php ..\..\pidocumenter\figaf\pidiff.php . old.xim new.xim 1
```
In this it is assumed that the two files old.xim and new.xim both are in the current folder, hence the . after pidiff.

After the script terminates it a fill called diff.xlsx or diff.xls is created in the selected directory.


# PI Documenter #
To use the PI documentation tool.
```
Usage:
 php   <pathto>pidocument.php <directory> <xim> <format> [<olddoc>]

<directory> Diretory where the files exists
<xim>       Filename of the xim file
<format>    Output format 1 for EXCEL2007 , 2 for EXCEL2003
<olddoc>    Old excel spreadsheet if the comments should be copied
```
The olddoc is optional, it should be an excel sheet genereted earlier with the tool. Comments is copied to the new document, where the paths matchs.


An example on this could be the following.
```
C:\Projects\test\doc>php ..\..\pidocumenter\figaf\pidocument.php . new.xim 1 oldcom.xlsx
```

After running the script a documentation file is created with the name of mapping.xlsx.