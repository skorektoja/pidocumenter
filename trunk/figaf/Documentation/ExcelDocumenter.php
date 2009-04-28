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

class ExcelDocumenter {
    /**
     * Returns $mappingInfo.
     * @see ExcelDocumenter::$mappingInfo
     */
    public function getMappingInfo()
    {
        return $this->mappingInfo;
    }
    
    /**
     * Sets $mappingInfo.
     * @param object $mappingInfo
     * @see ExcelDocumenter::$mappingInfo
     */
    public function setMappingInfo($mappingInfo)
    {
        $this->mappingInfo = $mappingInfo;
    }

    /**
     * Returns $excelObject.
     * @see ExcelDocumenter::$excelObject
     */
    public function getExcelObject()
    {
        return $this->excelObject;
    }

    /**
     * Returns $directory.
     * @see ExcelDocumenter::$directory
     */
    public function getDirectory()
    {
        return $this->directory;
    }
    
    /**
     * Sets $directory.
     * @param object $directory
     * @see ExcelDocumenter::$directory
     */
    public function setDirectory($directory)
    {
        $this->directory = $directory;
    }

    /**
     * Returns $inputfile.
     * @see ExcelDocumenter::$inputfile
     */
    public function getInputfile()
    {
        return $this->inputfile;
    }
    
    /**
     * Sets $inputfile.
     * @param object $inputfile
     * @see ExcelDocumenter::$inputfile
     */
    public function setInputfile($inputfile)
    {
        $this->inputfile = $inputfile;
    }

	/**
	 * Directory for where the files are stored
	 *
	 * @var String
	 */
	private $directory;

	/**
	 * Generate  Excel 2007 file if true else old format.
	 *
	 * @var int
	 */
	private $generateFormat;
	private $generateExcel2007;
	private $oldExcel2007;
	private $inputfile;
	
	/**
	 * The object where the data is written
	 * @var PHPExcel
	 */
	private $excelObject;

	const EXCEL2007   = 1;
	const EXCEL2003   = 2;
	const CSV         = 3;
	const HTML		  = 4;

   /**
    * Information about the current mapping 
    * @var ObjectInfo
    */
	private $mappingInfo;

	/**
	 * Constructor for the Excel documenter
	 *
	 * @param unknown_type $dir
	 * @param unknown_type $generateExcel2007
	 */
	public function __construct($dir, $inputfile, $generateFormat,$oldFormat){

		/** PHPExcel */
		include_once 'PHPExcel.php';

		/** PHPExcel_IOFactory */
		include_once 'PHPExcel/IOFactory.php';
		include_once 'PerformDocumentation.php';
		include_once 'Util/Extract.php';

		$this->directory = $dir;

		$this->generateFormat = $generateFormat;
		$this->oldExcel2007 = $oldFormat;
		$this->inputfile = $inputfile;

	}




	public function run(){

		Extract::extractXMI($this->directory, $this->inputfile);

		$this->mappingInfo = new ObjectInfo($this->directory.'/OBJECT_INFO.xml');
		$objPHPExcel = new PHPExcel();

		$objPHPExcel->getActiveSheet()->setTitle('Documentation');


		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$objPHPExcel->setActiveSheetIndex(0);

		//create header information

		$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Mapping name');
		$objPHPExcel->getActiveSheet()->setCellValue('B1', $this->mappingInfo->getValue('NAME'));

		$objPHPExcel->getActiveSheet()->setCellValue('A2', 'Namespace');
		$objPHPExcel->getActiveSheet()->setCellValue('B2', $this->mappingInfo->getValue('NAMESPACE'));

		$objPHPExcel->getActiveSheet()->setCellValue('A3', 'SWCV');
		$objPHPExcel->getActiveSheet()->setCellValue('B3', $this->mappingInfo->getValue('SWCVERSION'));
		$commentA3 = $objPHPExcel->getActiveSheet()->getComment('A3')->getText()->createText('Software component Version');


		$objPHPExcel->getActiveSheet()->setCellValue('A4', 'ObjectId');
		$objPHPExcel->getActiveSheet()->setCellValue('B4', $this->mappingInfo->getValue('OBJECTID'));

		$objPHPExcel->getActiveSheet()->setCellValue('A5', 'Changed');
		$objPHPExcel->getActiveSheet()->setCellValue('B5', $this->mappingInfo->getValue('CHANGED'));

		$objPHPExcel->getActiveSheet()->setCellValue('A6', 'Changed By');
		$objPHPExcel->getActiveSheet()->setCellValue('B6', $this->mappingInfo->getValue('CHANGEDBY'));

		$objPHPExcel->getActiveSheet()->setCellValue('A7', 'Description');
		$objPHPExcel->getActiveSheet()->mergeCells('B7:D7');
		//format header
		$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
		$objPHPExcel->getActiveSheet()->duplicateStyle( $objPHPExcel->getActiveSheet()->getStyle('A1'), 'A1:B7' );


		$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(18);
		$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(60);
		$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(90);
		$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(60);


		$objPHPExcel->getActiveSheet()->setCellValue('A9', 'ID');
		$objPHPExcel->getActiveSheet()->setCellValue('B9', 'Target');
		$objPHPExcel->getActiveSheet()->setCellValue('C9', 'Mapping');
		$objPHPExcel->getActiveSheet()->setCellValue('D9', 'Comment');

		$objPHPExcel->getActiveSheet()->getStyle('A9')->getFont()->setBold(true);


		$objPHPExcel->getActiveSheet()->getStyle('A9')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

		$objPHPExcel->getActiveSheet()->duplicateStyle( $objPHPExcel->getActiveSheet()->getStyle('A9'), 'B9:D9' );

		// comments from old document
		$oldCommentMap = array();
		//if old excel filename exists, add the comments.
		if(is_file($this->directory.'/old.xls')){
			//open the old excel document
			$objReader = PHPExcel_IOFactory::createReader($this->oldExcel2007  ?'Excel2007':'Excel5');

			$objExcelReader = $objReader->load($this->directory.'/old.xls');

			$objExcelReader->setActiveSheetIndex(0);



			for($i=10;$i<  $objExcelReader->getActiveSheet()->getHighestRow( ); $i++){

				$oldCommentMap[ $objExcelReader->getActiveSheet()->getCell("B".$i)->getValue()]
				= $objExcelReader->getActiveSheet()->getCell("D".$i)->getValue();

			}
		}



		// get mapping doc

		$metadata_string = file_get_contents($this->directory.'/metadata');
		// create the documentation mapping
		$xml  = new PerformDocumentation( $metadata_string, $objPHPExcel,$oldCommentMap);

		switch ($this->generateFormat){
			case self::EXCEL2007:
				$excelType = 'Excel2007';
				break;
			case self::EXCEL2003:
				$excelType = 'Excel5';
				break;
				 
			case self::CSV:
				$excelType = 'CSV';
				break;
			case self::HTML:
				$excelType = 'HTML';
				break;
			default:
				$excelType = 'Excel2007';
				break;

		}

		
		// write content
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $excelType);
		$objWriter->save($this->directory .'/output.xlsx');

       $this->excelObject = $objPHPExcel;

	}
}




/**
 * Simple parsing object. Parses information into a hash table, and data can be retrieved
 * with a key
 *
 */
class ObjectInfo{

	private $hashTabel;
	private $parser;
	private $value;

	/**
	 * Constructor
	 *
	 * @param unknown_type $filename name of the file to read XML form.
	 */
	public function  __construct($filename){
		$this->parser = xml_parser_create();
		xml_set_object($this->parser, $this);
		//        xml_parser_set_option($this->parser, XML_OPTION_CASE_FOLDING, false);
		xml_set_element_handler($this->parser, "start_element", "end_element");
		xml_set_character_data_handler($this->parser, "cdata");
		xml_parse($this->parser, file_get_contents($filename));

	}

	private function start_element($parser, $name, $attrs){
	}
	private  function cdata($parser, $cdata) {
		// save the current string
		$this->value = $cdata;
	}
	private function end_element ($parser, $name){
		$this->hashTabel[$name ] = $this->value;
	}
	/**
	 * Retrive information about a parameter
	 *
	 * @param unknown_type $name
	 */
	public function getValue($name){
		if (array_key_exists($name, $this->hashTabel))
		return $this->hashTabel[$name];
	}
}
?>