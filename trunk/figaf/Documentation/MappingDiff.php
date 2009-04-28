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
 /**
 * Writes the difference between two mappings.
 */

class MappingDifferance
{

    /**
     * @var documenter variable for create documentation
     */
    private $documenter;
    const STARTLINE = 10;

    /**
     * @var int describing which output format to use.
     */
    private $generateFormat;

    /**
     * Directory for where the files are stored
     *
     * @var String
     */
    private $directory;

    /**
     * The old filename
     * @var String
     */
    private $oldFileName="old.xim";
    /**
     * The new filename
     * @var String
     */
    private $newFileName = "new.xim";

	public function __construct($dir, $generateFormat, $oldFilename = "old.xim", $newFileName= "new.xim")
    {

      set_include_path(get_include_path().PATH_SEPARATOR.'..\\..\\lib\\'.PATH_SEPARATOR.'..\\');
        /** PHPExcel */
        include_once ("PHPExcel.php");


        /** PHPExcel_IOFactory */
        include_once 'PHPExcel/IOFactory.php';
        include_once 'ExcelDocumenter.php';

        $this->directory = $dir;
        $this->generateFormat = $generateFormat;
		$this->oldFileName = $oldFilename;
		$this->newFileName = $newFileName;

        $outputFormat = ExcelDocumenter::EXCEL2007;
        $oldFileExcel2007 = false;
        $this->documenter = new ExcelDocumenter($this->directory."/old", 'old.xim', $outputFormat, $oldFileExcel2007);


    }
    public function run()
    {

        // generate the documentation for the two mappings, and collect the other information needed to make the difference.
        list ($newMap, $newPath, $newMapInfo) = $this->prepareFile("new", $this->newFileName);
        list ($oldMap, $oldPath, $oldMapInfo) = $this->prepareFile("old", $this->oldFileName);


        $paths = $this->mergeInOrder($newPath, $oldPath);



        $objPHPExcel = new PHPExcel();


        // metadata
        $objPHPExcel->getProperties()->setCreator("figaf");
        $objPHPExcel->getProperties()->setLastModifiedBy("figaf");
        $objPHPExcel->getProperties()->setTitle("PI Mapping diff");
        $objPHPExcel->getProperties()->setKeywords("pi mapping diff");



        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        $objPHPExcel->getActiveSheet()->setTitle('Mapping difference');

        //insert information about the mappping
        $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Mapping name');
        $objPHPExcel->getActiveSheet()->setCellValue('C1', $oldMapInfo->getValue('NAME'));
        $objPHPExcel->getActiveSheet()->setCellValue('D1', $newMapInfo->getValue('NAME'));
  
        $objPHPExcel->getActiveSheet()->mergeCells('A1:B1');

 $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Namespace');
        $objPHPExcel->getActiveSheet()->setCellValue('C2', $oldMapInfo->getValue('NAMESPACE'));
        $objPHPExcel->getActiveSheet()->setCellValue('D2', $newMapInfo->getValue('NAMESPACE'));
        $objPHPExcel->getActiveSheet()->mergeCells('A2:B2');

        $objPHPExcel->getActiveSheet()->setCellValue('A3', 'SWCV');
        $objPHPExcel->getActiveSheet()->setCellValue('C3', $oldMapInfo->getValue('SWCVERSION'));
        $objPHPExcel->getActiveSheet()->setCellValue('D3', $newMapInfo->getValue('SWCVERSION'));
        $objPHPExcel->getActiveSheet()->mergeCells('A3:B3');

        $objPHPExcel->getActiveSheet()->setCellValue('A4', 'ObjectId');
        $objPHPExcel->getActiveSheet()->setCellValue('C4', $oldMapInfo->getValue('OBJECTID'));
        $objPHPExcel->getActiveSheet()->setCellValue('D4', $newMapInfo->getValue('OBJECTID'));
        $objPHPExcel->getActiveSheet()->mergeCells('A4:B4');

        $objPHPExcel->getActiveSheet()->setCellValue('A5', 'Changed');
        $objPHPExcel->getActiveSheet()->setCellValue('C5', $oldMapInfo->getValue('CHANGED'));
        $objPHPExcel->getActiveSheet()->setCellValue('D5', $newMapInfo->getValue('CHANGED'));
        $objPHPExcel->getActiveSheet()->mergeCells('A5:B5');

        $objPHPExcel->getActiveSheet()->setCellValue('A6', 'Changed By');
        $objPHPExcel->getActiveSheet()->setCellValue('C6', $oldMapInfo->getValue('CHANGEDBY'));
        $objPHPExcel->getActiveSheet()->setCellValue('D6', $newMapInfo->getValue('CHANGEDBY'));
        $objPHPExcel->getActiveSheet()->mergeCells('A6:B6');


        //format the top document
        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(60);
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(70);
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(70);

        $objPHPExcel->getActiveSheet()->setCellValue('A'.(self::STARTLINE-1), 'Changed');
        $objPHPExcel->getActiveSheet()->setCellValue('B'.(self::STARTLINE-1), 'Target');
        $objPHPExcel->getActiveSheet()->setCellValue('C'.(self::STARTLINE-1), 'Old mapping');
        $objPHPExcel->getActiveSheet()->setCellValue('D'.(self::STARTLINE-1), 'New Mapping');


        $objPHPExcel->getActiveSheet()->getStyle('A'.(self::STARTLINE-1))->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('A'.(self::STARTLINE-1))->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

//       $objPHPExcel->getActiveSheet()->duplicateStyle($objPHPExcel->getActiveSheet()->getStyle('A'.(self::STARTLINE-1)),
//       'B'.(self::STARTLINE-1).':D'.(self::STARTLINE-1));

/**
        $oldMapping;
        $newMapping;
        for ($i = 0; $i < count($paths); $i++)
        {
            $path = $paths[$i];

            $index = $i+self::STARTLINE;
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$index, $path);

            if (array_key_exists($path, $oldMap))
            {
                $oldMapping = $oldMap[$path];
            }
            else
            {
                $oldMapping = "";
            }
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$index, $oldMapping);
            if (array_key_exists($path, $newMap))
            {
                $newMapping = $newMap[$path];

            } else
            {
                $newMapping = "";

            }

            $objPHPExcel->getActiveSheet()->setCellValue('D'.$index, $newMapping);
            // test if the mappings are different
            $different = "yes";

            if ($newMapping instanceof PHPExcel_RichText && $oldMapping instanceof PHPExcel_RichText)
            {

                if ($newMapping->getHashCode() == $oldMapping->getHashCode())
                {
                    $different = "no";
                }
            }

            if (!$newMapping instanceof PHPExcel_RichText && !$oldMapping instanceof PHPExcel_RichText)
            {

                if ($newMapping == $oldMapping)
                {
                    $different = "no";
                }
            }


            $objPHPExcel->getActiveSheet()->setCellValue('A'.$index, $different);


            //format current line
            $objPHPExcel->getActiveSheet()->getStyle('A'.$index)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

            $objPHPExcel->getActiveSheet()->getStyle('B'.$index)->getAlignment()->setWrapText(true);
            $objPHPExcel->getActiveSheet()->getStyle('B'.$index)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

            $objPHPExcel->getActiveSheet()->getStyle('C'.$index)->getAlignment()->setWrapText(true);
            $objPHPExcel->getActiveSheet()->getStyle('C'.$index)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
            $objPHPExcel->getActiveSheet()->getStyle('D'.$index)->getAlignment()->setWrapText(true);
            $objPHPExcel->getActiveSheet()->getStyle('D'.$index)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);


        }
		*/
		
		$index=12;
        // create filter
 //      $objPHPExcel->getActiveSheet()->setAutoFilter('A'.(self::STARTLINE-1).':D'.$index);

        // determinate the output type.
        $excelType = 'Excel2007';
        switch($this->generateFormat)
        {
            case ExcelDocumenter::EXCEL2007:
                $excelType = 'Excel2007';
                break;
            case ExcelDocumenter::EXCEL2003:
                $excelType = 'Excel5';
                break;

            case ExcelDocumenter::CSV:
                $excelType = 'CSV';
                break;
            case ExcelDocumenter::HTML:
                $excelType = 'HTML';
                break;

        }


        // write content
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $excelType);
        $objWriter->save($this->directory.'/diff.xlsx');

    }


    private function prepareFile($prefix,$filename)
    {
		
        $pathArray = array ();
        $mapArray = array ();

        mkdir($this->directory."/$prefix");

        copy($this->directory."/$filename", $this->directory."/$prefix/$filename");
        $this->documenter->setDirectory($this->directory."/$prefix");
        $this->documenter->setInputfile("$filename");
        $this->documenter->run();

        $objExcelReaderOld = $this->documenter->getExcelObject();
        $objExcelReaderOld->setActiveSheetIndex(0);

        for ($i = 10; $i < $objExcelReaderOld->getActiveSheet()->getHighestRow(); $i++)
        {
            $cell = $objExcelReaderOld->getActiveSheet()->getCell("B".$i);
            $path = ($cell->getValue() instanceof PHPExcel_RichText)?$cell->getValue()->getPlainText():
                $cell->getValue();

                $mapArray[$path] = $objExcelReaderOld->getActiveSheet()->getCell("C".$i)->getValue();
                $pathArray[] = $path;
            }
            return array ($mapArray, $pathArray, $this->documenter->getMappingInfo());
        }



/**
 * Create an intereplace where the two arrays is combined. 
 * Found at // http://dk2.php.net/manual/en/function.array-merge.php#73584
 * @return 
 * @param object $a  Master array
 * @param object $b   append thise at the end of the document
 */
private function mergeInOrder($a, $b) {

  for($i =0; $i<count($b);$i++){
  	$found = false;
	 for($j =0; $j<count($a);$j++){
	  	if($a[$j]== $b[$i]){
	  		$found = true;
			break;
	  	}
  	}
	if(!$found){
		array_push($a,$b[$i]);
	}
  }
 
  return $a;


} 

    }


	

?>
