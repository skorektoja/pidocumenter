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
 * @version    0.9 , 2009-04-28
 */


if ($argc != 5 || in_array($argv[1], array('--help', '-help', '-h', '-?'))) {
?>

This the pi diff tool.

  Usage:
  <?php echo $argv[0]; ?> <directory> <old> <new> <format> 

  <directory> Diretory where the files exists
  <old>       Filename of the old xim file
  <new>       Filename of the new xim file
  <format>    Output format 1 for EXCEL2007 , 2 for EXCEL2003
  
<?php  
 exit;
}

	$scriptdir=  substr( $argv[0],  0, strripos($argv[0], "\\")); 

	set_include_path(get_include_path() . PATH_SEPARATOR . "$scriptdir\\..\\lib\\". PATH_SEPARATOR . "$scriptdir");
	 include_once 'Documentation/MappingDiff.php';
	   
	   
	 $dirname= $argv[1];
	 $oldFilename= $argv[2];
	 $newFilename = $argv[3];
	 $format = $argv[4];
	 
	   

	   
	 $mapDiff= new MappingDifferance($dirname, $format , $oldFilename, $newFilename	);
 	$mapDiff->run();
	
	// rename the diff file to .xls if it is a excel2003 selected
	if($format == 2){
		exec ("mv $dirname\\diff.xlsx $dirname\\diff.xls ");
	}
	
	
	// delete temp files
//	exec("rd /S /Q $dirname\\new $dirname\\old");
	


?>