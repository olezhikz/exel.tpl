<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2014 PHPExcel
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
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.8.0, 2014-03-02
 */

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
date_default_timezone_set('Europe/London');

header('Content-Type: text/html; charset=utf-8');

/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';

$inputFileName = 'test.xls';
$objReader = new PHPExcel_Reader_Excel5();
$objPHPExcel = $objReader->load($inputFileName);
$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);

$maxNum = 11;
$resultColumn = 'N';
//$abc 		= array('C','D','E','F','G');
$abc = array(
	'C' => 'H',
	'D' => 'I',
	'E' => 'J',
	'F' => 'K',
	'G' => 'L'
	);
$sum 		=  array();
$numArr = array(
	1 => '1',
	2 => '0.9',
	3 => '0.8',
	4 => '0.7',
	5 => '0.6',
	6 => '0.5',
	7 => '0.4',
	8 => '0.3',
	9 => '0.2',
	10=> '0.1'
	);

for ($i=1; $i< count($sheetData) ; $i++) {
	foreach ($sheetData[$i] as $key => $cell) {
		foreach ($abc as $letterID => $latterValue) {
			if ($key == $letterID) {
				if (preg_match("/((\S+,){1,})/s", $sheetData[$i][$key])) {
					$resArr =  explode(",", $sheetData[$i][$key]);
					foreach ($resArr as $id => $num) {
						foreach ($numArr as $checKey => $checkValue) {
							if ($num == $checKey ) {
								$resArr[$id] = $checkValue;
							}
						}
					}
					$sheetData[$i][$key] = array_sum($resArr);
				} else{
					$checkArr =  explode(" ", $sheetData[$i][$key]);
					$checkMaxNum = (int)$checkArr[0];
					if( $checkMaxNum <= $maxNum && $checkMaxNum != $maxNum ){
						foreach ($numArr as $checKey2 => $checkValue2) {
							if ($checkMaxNum == $checKey2 ) {
								 $checkMaxNum = $checkValue2;
							}
						}
						$sheetData[$i][$key]  = $checkMaxNum;
					} else {
						$sheetData[$i][$key]  = 0;
					}
				}
			}
			$sum[$i][$letterID] = $sheetData[$i][$letterID] * $sheetData[$i][$latterValue];
			$sheetData[$i][$resultColumn] = array_sum($sum[$i]);
		}
	}
}
// echo "<pre>";
//var_dump($sheetData);
// echo "</pre>";

// Create new PHPExcel object
$objWriter = new PHPExcel();

for ($i=1; $i< count($sheetData); $i++) {
	foreach ($sheetData[$i] as $latter => $cellValue) {
		$objWriter->setActiveSheetIndex(0)->setCellValue( $latter.$i, $cellValue);
	}
}

// Rename worksheet
$objWriter->getActiveSheet()->setTitle('DATA');

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objWriter->setActiveSheetIndex(0);

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="01simple.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

$out = PHPExcel_IOFactory::createWriter($objWriter, 'Excel5');
$out->save('php://output');
exit;


