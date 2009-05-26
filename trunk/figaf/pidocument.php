<?php
/**
 * pidocumenter
 *
 * Copyright (c) 2008 - 2009 pidocumenter
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   pidocumenter
 * @package    pidocumenter
 * @copyright  Copyright (c) 2008 - 2009 pidocumenter (http://www.figaf.com/services/pi-documenter.html)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    0.9.1 , 2009-05-26
 */
if ($argc < 4 || $argc > 5 || in_array($argv[1], array ('--help', '-help', '-h', '-?')))
{
?>
This the pi documentation tool.
  Usage:
  <?php echo $argv[0]; ?> <directory> <xim> <format> [<olddoc>]

  <directory> Diretory where the files exists
  <xim>       Filename of the xim file
  <format>    Output format 1 for EXCEL2007 , 2 for EXCEL2003
  <olddoc>    Old excel spreadsheet if the comments should be copied
  
<?php
exit ;
}

$scriptdir = substr($argv[0], 0, strripos($argv[0], "\\"));

set_include_path(get_include_path().PATH_SEPARATOR."$scriptdir\\..\\lib\\".PATH_SEPARATOR."$scriptdir");

include_once 'Documentation/ExcelDocumenter.php';
$dirname = $argv[1];
$ximname = $argv[2];
$format = $argv[3];

mkdir  ("$dirname\\tempdoc");
copy("$dirname\\$ximname", "$dirname\\tempdoc\\$ximname");

$oldExcelformat= ExcelDocumenter::EXCEL2007;
if ($argc == 5)
{
    $oldFilename= $argv[4];
	copy("$dirname\\$oldFilename", "$dirname\\tempdoc\\old.xls");
	if (stristr ($oldFilename,".xlsx")>0){
			$oldExcelformat = 	ExcelDocumenter::EXCEL2007	;
	}else{
			$oldExcelformat = 	ExcelDocumenter::EXCEL2003	;
	}

}

// create tempory directories



$documenter = new ExcelDocumenter("$dirname\\tempdoc", $ximname, $format, $oldExcelformat);
$documenter->run();
$mappingName = $documenter->getMappingInfo()->getValue('NAME');
if ($format == ExcelDocumenter::EXCEL2007)
{
    // excel 2007
	unlink  ($dirname.'\\'.$mappingName.'.xlsx');
    rename($dirname.'\\tempdoc\\output.xlsx', $dirname.'\\'.$mappingName.'.xlsx');
} else
{
    // excel 2005
	unlink  ($dirname.'\\'.$mappingName.'.xls');
    rename($dirname.'\\tempdoc\\output.xlsx', $dirname.'\\'.$mappingName.'.xls');
}


// delete temp files
exec("rd /S /Q $dirname\\tempdoc ");

?>
