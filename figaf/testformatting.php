<?php
 set_include_path(get_include_path().PATH_SEPARATOR.'..\\lib\\');
        /** PHPExcel */
        include_once ("PHPExcel.php");


        /** PHPExcel_IOFactory */

        include_once 'PHPExcel/IOFactory.php';
	
	$objPHPExcel = new PHPExcel();
// Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        $objPHPExcel->getActiveSheet()->setTitle('Mapping difference');

        //insert information about the mappping
        $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Mapping name');
        $objPHPExcel->getActiveSheet()->setCellValue('C1', 'name');

        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
		
		     // write content
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('test.xlsx');
  	
		
?>