<?php

header( "Content-type: text/html; charset=utf-8" );

$inputFileType = 'Excel5';
$inputFileName = '2013.xls';

/** PHPExcel_IOFactory */
include 'Classes/PHPExcel/IOFactory.php';

define( 'EOL', ( PHP_SAPI == 'cli' ) ? PHP_EOL : '<br />' );
/**
 * Create a new Reader of the type defined in $inputFileType *
 */
$objReader = PHPExcel_IOFactory::createReader( $inputFileType );


/**
 * Load $inputFileName to a PHPExcel Object *
 */
$objPHPExcel = $objReader->load( $inputFileName );
echo '<hr />';
echo 'Reading the number of Worksheets in the WorkBook<br />';


/**
 * Use the PHPExcel object's getSheetCount() method to get a count of the number
 * of WorkSheets in the WorkBook
 */

$sheetCount = $objPHPExcel->getSheetCount();
echo 'There ', ( ( $sheetCount == 1 ) ? 'is' : 'are' ), ' ', $sheetCount, ' WorkSheet', ( ( $sheetCount == 1 ) ? '' : 's' ), ' in the WorkBook<br /><br />';
echo 'Reading the names of Worksheets in the WorkBook<br />';


/**
 * Use the PHPExcel object's getSheetNames() method to get an array listing the
 * names/titles of the WorkSheets in the WorkBook
 */
$sheetNames = $objPHPExcel->getSheetNames();

var_dump( $sheetNames );
foreach ( $sheetNames as $sheetIndex => $sheetName )
{
	echo 'WorkSheet #', $sheetIndex, ' is named "', $sheetName, '"<br />';
}

// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;



echo '<hr />';

/**  Read the document's creator property  **/
$creator = $objPHPExcel->getProperties()->getCreator();
echo '<b>Document Creator: </b>',$creator,'<br />';

/**  Read the Date when the workbook was created (as a PHP timestamp value)  **/
$creationDatestamp = $objPHPExcel->getProperties()->getCreated();
/**  Format the date and time using the standard PHP date() function  **/
$creationDate = date('l, d<\s\up>S</\s\up> F Y',$creationDatestamp);
$creationTime = date('g:i A',$creationDatestamp);
echo '<b>Created On: </b>',$creationDate,' at ',$creationTime,'<br />';

/**  Read the name of the last person to modify this workbook  **/
$modifiedBy = $objPHPExcel->getProperties()->getLastModifiedBy();
echo '<b>Last Modified By: </b>',$modifiedBy,'<br />';

/**  Read the Date when the workbook was last modified (as a PHP timestamp value)  **/
$modifiedDatestamp = $objPHPExcel->getProperties()->getModified();
/**  Format the date and time using the standard PHP date() function  **/
$modifiedDate = date('l, d<\s\up>S</\s\up> F Y',$modifiedDatestamp);
$modifiedTime = date('g:i A',$modifiedDatestamp);
echo '<b>Last Modified On: </b>',$modifiedDate,' at ',$modifiedTime,'<br />';

/**  Read the workbook title property  **/
$workbookTitle = $objPHPExcel->getProperties()->getTitle();
echo '<b>Title: </b>',$workbookTitle,'<br />';

/**  Read the workbook description property  **/
$description = $objPHPExcel->getProperties()->getDescription();
echo '<b>Description: </b>',$description,'<br />';

/**  Read the workbook subject property  **/
$subject = $objPHPExcel->getProperties()->getSubject();
echo '<b>Subject: </b>',$subject,'<br />';

/**  Read the workbook keywords property  **/
$keywords = $objPHPExcel->getProperties()->getKeywords();
echo '<b>Keywords: </b>',$keywords,'<br />';

/**  Read the workbook category property  **/
$category = $objPHPExcel->getProperties()->getCategory();
echo '<b>Category: </b>',$category,'<br />';

/**  Read the workbook company property  **/
$company = $objPHPExcel->getProperties()->getCompany();
echo '<b>Company: </b>',$company,'<br />';

/**  Read the workbook manager property  **/
$manager = $objPHPExcel->getProperties()->getManager();
echo '<b>Manager: </b>',$manager,'<br />';






header( "Content-type: text/html; charset=utf-8" );

/** PHPExcel_IOFactory */
include 'Classes/PHPExcel/IOFactory.php';

/**
 * Error reporting
 */
error_reporting( E_ALL );
ini_set( 'display_errors', TRUE );
ini_set( 'display_startup_errors', TRUE );
date_default_timezone_set( 'Europe/London' );
define( 'EOL', ( PHP_SAPI == 'cli' ) ? PHP_EOL : '<br />' );

$file_name = '2013.xls';

/**
 * Include PHPExcel_IOFactory
 */
require_once 'Classes/PHPExcel/IOFactory.php';
if ( !file_exists( $file_name ) )
{
	exit( "The 2013.xls file is not exists.\n" );
}

echo date( 'H:i:s' ), " Load workbook from 2013.xls file", EOL;


$callStartTime = microtime( true );
// $objPHPExcel = PHPExcel_IOFactory::load( $file_name );


$objReader = new PHPExcel_Reader_Excel5();
// $objReader->setLoadSheetsOnly( );
$objPHPExcel = $objReader->load( $file_name );

$callEndTime = microtime( true );
$callTime = $callEndTime - $callStartTime;
echo 'Call time to load Workbook was ', sprintf( '%.4f', $callTime ), " seconds", EOL;

// Echo memory usage
echo date( 'H:i:s' ), ' Current memory usage: ', ( memory_get_usage( true ) / 1024 / 1024 ), " MB", EOL;


/**
 * Use the PHPExcel object's getSheetNames() method to get an array listing the
 * names/titles of the WorkSheets in the WorkBook
 */
$sheetNames = $objPHPExcel->getSheetNames();

echo 'Sheet Count:';$objPHPExcel->getSheetCount() . EOL;

//$objPHPExcel->getSheet( 9 )->getStyle('K10')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
//$objPHPExcel->getSheet( 9 )->getStyle('K10')->set
//$objPHPExcel->getSheet( 0 )->setCellValue( 'A1', 'Why');
//$objPHPExcel->getSheet( 1 )->setCellValue( 'A1', 'All');

// echo 'C, 26:' . EOL;
// echo $objPHPExcel->getSheet( 0 )->getCellByColumnAndRow( 'C', 26 ) . EOL;

// echo 'AV, 4:' . EOL;
// echo $objPHPExcel->getSheet( 0 )->getCellByColumnAndRow( 'AV', 4 ) . EOL;

//echo $objPHPExcel->getSheet( 0 )->getCell( 'B20' )->getValue() . EOL;

$objPHPExcel->setActiveSheetIndex( 0 );
//echo 'sheet1:' . $objPHPExcel->getSheet( 0 )->setCellValue( 'C20', 1000, true ) . EOL;
//echo 'sheet2:' . $objPHPExcel->getSheet( 1 )->setCellValue( 'C20', 1000, true ) . EOL;

$objClonedWorksheet = clone $objPHPExcel->getSheetByName( $objPHPExcel->getActiveSheet()->getTitle() );
$objClonedWorksheet->setTitle( 'Copy of2' . $objPHPExcel->getActiveSheet()->getTitle() );
$objPHPExcel->addSheet( $objClonedWorksheet );


/* echo 'sheet3:' . $objPHPExcel->getSheet( 2 )->setCellValue( 'C20', 1000, true ) . EOL;
 echo 'sheet4:' . $objPHPExcel->getSheet( 3 )->setCellValue( 'C20', 1000, true ) . EOL;
echo 'sheet5:' . $objPHPExcel->getSheet( 4 )->setCellValue( 'C20', 1000, true ) . EOL;
echo 'sheet6:' . $objPHPExcel->getSheet( 5 )->setCellValue( 'C20', 1000, true ) . EOL;
echo 'sheet7:' . $objPHPExcel->getSheet( 6 )->setCellValue( 'C20', 1000, true ) . EOL;
echo 'sheet8:' . $objPHPExcel->getSheet( 7 )->setCellValue( 'C20', 1000, true ) . EOL;
echo 'sheet9:' . $objPHPExcel->getSheet( 9 )->setCellValue( 'C20', 1000, true ) . EOL;
echo 'sheet10:' . $objPHPExcel->getSheet( 10 )->setCellValue( 'C20', 1000, true ) . EOL; */

// echo $objPHPExcel->getSheet( 0 )->getCellByColumnAndRow( 13, 4 )->getValue() . EOL;


// //print_r( $objPHPExcel->getSheet( 0 ) );

foreach ( $sheetNames as $sheetIndex => $sheetName )
{
	echo $sheetIndex . EOL;
	$objPHPExcel->setActiveSheetIndex( $sheetIndex );

	$activeSheet = $objPHPExcel->getActiveSheet();

	// Set cell number formats
	echo date('H:i:s') , $objPHPExcel->getProperties()->getTitle() , " Set cell number formats" , EOL;
	echo $objPHPExcel->getActiveSheet()->setCellValue( 'C20', 1000, true ) . EOL;

	//var_dump( $activeSheet );
	// 	echo 'WorkSheet #', $sheetIndex, ' is named "', $sheetName, '"<br />';
}



$objWriter = new PHPExcel_Writer_Excel5( $objPHPExcel );
$objWriter->setPreCalculateFormulas( false );
$objWriter->save( 'new.xls' );



/* $objWriter = PHPExcel_IOFactory::createWriter( $objPHPExcel, 'Excel5' );
 $objWriter->save( $objPHPExcel->getProperties()->getTitle() . 'new' . '.xls' );
*/



