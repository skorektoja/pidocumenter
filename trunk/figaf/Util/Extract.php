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

/**
 * Extract the XMI file to the directory, where the file is stored. 
 *
 */

class Extract{
	
	
	public static function extractXMI($directory, $filename){
		$zip = new ZipArchive;
		$res = $zip->open($directory.'/'.$filename);

		
			
		if ($res === TRUE) {
			$zip->extractTo($directory , array ('ZIPPED_META_DATA.zip','OBJECT_INFO.xml'));
			$zip->close();

		} else {
			//try using unix unzip because a bug when exporting from Citrix. 
			exec("unzip  $directory/$filename -d $directory");
		}

		$res = $zip->open($directory.'/ZIPPED_META_DATA.zip');

		if ($res === TRUE) {
			$zip->extractTo($directory);
			$zip->close();
		} else {
			echo 'failed';
	              exit;
		}
			
			

	
	
	}
}
?>