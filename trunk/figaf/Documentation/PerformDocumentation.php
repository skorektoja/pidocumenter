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

class PerformDocumentation {
	private $parser;
	//    private $objPHPExcel;

	/**
	* Which pi version is used. 
	*/
	private $piversion = "";

	private $inTransformation = false;
	private $inFunctionStorage = false;
	private $index = 0;
	private $currentPath = "";
	private $currentMapping = "";

	private $linenumber = 9;
	private $docPHPExcel = null;
	private $oldDocMap = null;
	private $inFuncParameters = false;
	private $parameterString = "";
	private $currentString = "";
	private $mapRichText;
	/**
	 * Placeholder for function parameters.
	 *
	 * @var unknown_type
	 */
	private $functionParams;
	private $currentFunctionName;

 /**
  * Array with information about userdefined functions. 
  */
	private $udfMap = array ();
	private $udfParameters = "";
	private $udfIndex = 0;

	private $hiddenMap = array ();

	/**
	 * Name of the current attribute
	 *
	 * @var unknown_type
	 */
	private $currentParameterName;

	public function __construct($data, $objPHPExcel, $oldCommentMap) {

		$this->docPHPExcel = $objPHPExcel;
		$this->oldDocMap = $oldCommentMap;
		$this->parser = xml_parser_create();
		// Change & to allow for parsing of data in CDATA.
		$dataRewritten  = str_replace ("&", "&amp;", $data);
		xml_set_object($this->parser, $this);
		//        xml_parser_set_option($this->parser, XML_OPTION_CASE_FOLDING, false);
		xml_set_element_handler($this->parser, "start_element", "end_element");
		xml_set_character_data_handler($this->parser, "cdata");
		xml_parse($this->parser, $dataRewritten);

	}

	private function formatCurrentLine() {
		// write content
		$this->docPHPExcel->getActiveSheet()->setCellValue('B' . $this->linenumber, $this->currentPath);

		// format
		$this->docPHPExcel->getActiveSheet()->getStyle('C' . $this->linenumber)->getAlignment()->setWrapText(true);
		$this->docPHPExcel->getActiveSheet()->getStyle('B' . $this->linenumber)->getAlignment()->setVertical(PHPExcel_Style_Alignment :: VERTICAL_TOP);
		$this->docPHPExcel->getActiveSheet()->getStyle('C' . $this->linenumber)->getAlignment()->setVertical(PHPExcel_Style_Alignment :: VERTICAL_TOP);

		// hide row if target has been set to hide. 
		foreach ($this->hiddenMap as $hiddenNode) {
			$pos = strpos($this->currentPath, $hiddenNode . '/');
			if (!($pos === false) or $this->currentPath == $hiddenNode) {
				$this->docPHPExcel->getActiveSheet()->getRowDimension($this->linenumber)->setVisible(false);
			}
		}

		// Find old comment
		if (array_key_exists($this->currentPath, $this->oldDocMap)) {
			$comment = $this->oldDocMap[$this->currentPath];
			$this->docPHPExcel->getActiveSheet()->setCellValue('D' . $this->linenumber, $comment);
		}
	}

	private function start_element($parser, $name, $attrs) {

		if (($name == 'MAPPINGTOOL' or $name == 'PROJECT') and array_key_exists('VERSION', $attrs)) {

			$this->piversion = $attrs['VERSION'];
		}
		if ($name == 'TRANSFORMATION') {

			$this->inTransformation = true;
		}
		if ($name == 'FUNCTIONSTORAGE') {

			$this->inFunctionStorage = true;
		}

		if ($this->inTransformation) {
			//	echo "$name \n";
			if ($name == 'BRICK' and $attrs['TYPE'] == 'Dst') {

				if ($this->linenumber >= 10) {
					$this->formatCurrentLine();
					//if no data exists for an element create a empty string
					if(strlen($this->mapRichText->getPlainText())==0)
						$this->mapRichText->createText(' ');
				
				}
				$this->linenumber++;
				
				$this->mapRichText = new PHPExcel_RichText($this->docPHPExcel->getActiveSheet()->getCell('C' . $this->linenumber));

				$this->currentPath = $attrs['PATH'];
				$this->index = 0;

			} else
				if ($name == 'BRICK' and $attrs['TYPE'] == 'Src') {

					if (array_key_exists('CONTEXT', $attrs)) {

						$this->mapRichText->createText($this->getIndentation($this->index));
						$contextName = $this->mapRichText->createTextRun($attrs['CONTEXT']);
						$contextName->getFont()->setBold(true);
						$this->mapRichText->createText(substr($attrs['PATH'], strlen($attrs['CONTEXT'])));
					} else {
						$this->mapRichText->createText($this->getIndentation($this->index) . $attrs['PATH']);
					}
					$this->index++;

				} else
					if ($name == 'BRICK' and $attrs['TYPE'] == 'Func') {
						$this->mapRichText->createText($this->getIndentation($this->index));
						$functionName = $this->mapRichText->createTextRun($attrs['FNAME']);
						$functionName->getFont()->setBold(true);

						$this->index++;
						//create placeholder for parameters
						$this->functionParams[$this->index] = $this->mapRichText->createTextRun(' ');
						$this->currentFunctionName = $attrs['FNAME'];

					} else
						if ($name == 'PROPERTY' and $attrs['NAME'] == 'switchedOff') {
							// current node is hidden
							// default create a text element otherwhice this node could cause problems. 
							$this->mapRichText->createTextRun(' ');

							array_push($this->hiddenMap, $this->currentPath);
							$this->inFuncParameters = false;
							$this->parameterString = '';
						} else
							if ($name == 'PARAMETER' or $name == 'BINDINGS') {
								// in parameters for a function

								$this->inFuncParameters = true;
								$this->parameterString = '';

							} else
								if ($this->inFuncParameters == true and $this->piversion == 'XI7.1') {
									// we are defining parameters
									if (array_key_exists('NAME', $attrs)) {
										$this->currentParameterName = $attrs['NAME'];
									}
								} else
									if ($this->inFuncParameters == true) {
										// we are defining parameters
										$this->parameterString .= (strlen($this->currentString) > 0 ? $this->currentString . ' ' : '') . $name . '=';
										$this->currentString = "";

									}

		}

		// create array for 
		if ($this->inFunctionStorage) {
			// update the paramterstring string
			if ($name == "ARGUMENT") {
				if (strlen($this->udfParameters) > 0)
					$this->udfParameters .= ', ';
				$this->udfParameters .= $attrs['JTP'] . ' ' . $attrs['NM'];
			}

		}

	}

	private function end_element($parser, $name) {
		if ($name == 'BRICK') {
			$this->index--;
		}

		// end the parameter gatering
		if ($this->inFuncParameters and $this->piversion == 'XI7.1' and ($name == 'VALUE' or $name == 'PROPERTY')) {
			if (strlen($this->currentString) > 0 or strlen($this->currentParameterName) > 0) {
				$this->parameterString .= (strlen($this->parameterString) > 0 ? ' ' : '') . $this->currentParameterName . '=' . $this->currentString;
			}
			$this->currentString = "";
			$this->currentParameterName = "";

		} else
			if ($name == 'BINDINGS' and $this->inFuncParameters and $this->piversion == 'XI7.1') {
				// end the parameter gatering for PI 7.1
				$this->inFuncParameters = false;

				$this->functionParams[$this->index]->setText(' ' . $this->parameterString . $this->currentString);

				$this->currentString = "";
				$this->parameterString = "";
			} else
				if ($name == 'PARAMETER' and $this->inFuncParameters) {
					$this->inFuncParameters = false;
					if ($this->currentFunctionName == 'const') {
						$this->mapRichText->createText(' = \'' . $this->currentString . '\'');
					} else {
						$this->mapRichText->createText($this->getIndentation($this->index) . $this->parameterString . $this->currentString);
					}

					$this->currentString = "";
					$this->parameterString = "";
				}

		// Update table with udf
		if ($this->inFunctionStorage) {
			if ($name == 'SIGNATURE') {
				$this->udfMap[$this->udfIndex]['SIGNATURE'] = $this->udfParameters;
				$this->udfParameters = "";
			} else if($name=='FUNCTIONSTORAGE'){
				$this->inFunctionStorage = false;
			}
			else
			 // Java code
				if ($name == 'TEXT') {
					$this->udfMap[$this->udfIndex]['CODE'] =($this->currentString);
					$this->currentString="";
				} else
					if ($name == 'FUNCTIONMODEL') { // the current function has been found
						$this->udfIndex++;
					} else {
						$this->udfMap[$this->udfIndex][$name] = $this->currentString;
					}
				$this->currentString = "";

		}

		if ($name == 'TRANSFORMATION') {

			$this->inTransformation = false;
			$this->formatCurrentLine();
		}
		if ($name == 'FUNCTIONSTORAGE') {
			$this->inFunctionStorage = false;
			// create the UDF functions
			$this->docPHPExcel->createSheet();
			$this->docPHPExcel->setActiveSheetIndex(1);
			$this->docPHPExcel->getActiveSheet()->setTitle('UDF');
			$this->docPHPExcel->getActiveSheet()->setCellValue('A1', "Name");
			$this->docPHPExcel->getActiveSheet()->setCellValue('B1', "Signature");
			$this->docPHPExcel->getActiveSheet()->setCellValue('C1', "Code");
			$this->docPHPExcel->getActiveSheet()->setCellValue('D1', "Type");


		$this->docPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(18);
		$this->docPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(60);
		$this->docPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(90);
		$this->docPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(4);
		
		
		$this->docPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
		$this->docPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
		$this->docPHPExcel->getActiveSheet()->duplicateStyle( $this->docPHPExcel->getActiveSheet()->getStyle('A1'), 'A1:D1' );
		
		  
			for ($i = 0; $i < count($this->udfMap); $i++) {
				$this->docPHPExcel->getActiveSheet()->setCellValue('A' . ($i +2), $this->udfMap[$i]["NAME"]);
				$this->docPHPExcel->getActiveSheet()->setCellValue('B' . ($i +2), $this->udfMap[$i]["SIGNATURE"]);
				$this->docPHPExcel->getActiveSheet()->setCellValue('C' . ($i +2), $this->udfMap[$i]["CODE"]);
				$this->docPHPExcel->getActiveSheet()->setCellValue('D' . ($i +2), $this->udfMap[$i]["TYPE"]);
				// formattin for the line
				$this->docPHPExcel->getActiveSheet()->getStyle('C' . ($i +2))->getAlignment()->setWrapText(true);
		$this->docPHPExcel->getActiveSheet()->getStyle('A' . ($i +2))->getAlignment()->setVertical(PHPExcel_Style_Alignment :: VERTICAL_TOP);
		$this->docPHPExcel->getActiveSheet()->getStyle('B' . ($i +2))->getAlignment()->setVertical(PHPExcel_Style_Alignment :: VERTICAL_TOP);
		$this->docPHPExcel->getActiveSheet()->getStyle('C' . ($i +2))->getAlignment()->setVertical(PHPExcel_Style_Alignment :: VERTICAL_TOP);
		$this->docPHPExcel->getActiveSheet()->getStyle('D' .  ($i +2))->getAlignment()->setVertical(PHPExcel_Style_Alignment :: VERTICAL_TOP);
		

			}

			$this->docPHPExcel->setActiveSheetIndex(0);
			
		}

	}

	private function cdata($parser, $cdata) {
		// save the current string
		if ($this->inFuncParameters ) {
			$this->currentString = $cdata;
		}
		// functions can contain multiply parse elements, therefore
		// do they need to be concated
		if($this->inFunctionStorage){
			$this->currentString .= $cdata;
		}
	}
	/**
	 * Get a number of indentations
	 *
	 */
	private function getIndentation($no) {
		$indendation = '';
		if ($no > 0)
			$indendation = "\n";
		for ($i = 0; $i < $no; $i++) {
			$indendation .= '  ';
		}
		return $indendation;
	}
	
	public function getUdfMap(){
		return $this->udfMap;
	}
	

}
?>