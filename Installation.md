# Introduction #

How to install the pidocument tool on you computer


# Steps #

  1. Download and install PHP from http://www.php.net/downloads.php
  1. Download the pidocumenter release from the download page
  1. Extract the pidocumenter zip file to a directory
  1. Download PHPExcel from [codeplex](http://phpexcel.codeplex.com/Release/ProjectReleases.aspx?ReleaseId=10715). Notice there is currently a problem with formatting on diff tool with release 1.6.7 so please use 1.6.6 instead.
  1. Extract the PHPExcel. Copy the files from the zip files folder classes to the lib in the pidocument folder. It is the file PHPExcel.php and the folder (with subfolders) PHPExcel.
  1. You are now able to run the scripts.
  1. The build in unzip tool does not work in some instance. I have therefore used the unzip from http://sourceforge.net/projects/unxutils/. Download the zip file, extract the file and update your path to include `<extractfolder>` \usr\local\wbin