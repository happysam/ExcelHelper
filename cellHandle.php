<?php
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

$file_name = '2013.xls';

/**
 * Include PHPExcel_IOFactory
 */
require_once 'Classes/PHPExcel/IOFactory.php';
if ( !file_exists( $file_name ) )
{
	exit( 'The ' . $file_name . " file is not exists.\n" );
}

echo date( 'H:i:s' ), ' Load workbook from ' . $file_name . ' file', EOL;

$callStartTime = microtime( true );

$objReader = new PHPExcel_Reader_Excel5();
$objPHPExcel = $objReader->load( $file_name );

$callEndTime = microtime( true );
$callTime = $callEndTime - $callStartTime;
echo 'Call time to load Workbook was ', sprintf( '%.4f', $callTime ), " seconds", EOL , EOL;

// Echo memory usage
echo date( 'H:i:s' ), ' Current memory usage: ', ( memory_get_usage( true ) / 1024 / 1024 ), " MB", EOL;

/**
 * Use the PHPExcel object's getSheetNames() method to get an array listing the
 * names/titles of the WorkSheets in the WorkBook
 */
$sheetNames = $objPHPExcel->getSheetNames();
echo 'Sheet Count:' . $objPHPExcel->getSheetCount() . EOL;

$year = idate( 'Y' );
foreach ( $sheetNames as $sheetIndex => $sheetName )
{
	$beginday_timestamp = gmmktime( 0, 0, 0, $sheetIndex + 1, 1, $year );  // 当月开始日时间戳
	$current_days = idate( 't', $beginday_timestamp );  //当月天数
	
	$finishday_timestamp = gmmktime( 0, 0, 0, $sheetIndex + 1, $current_days, $year );  // 当月结束日时间戳
	
	echo 'processing ' . ( $sheetIndex + 1 ) . ' Sheet' . EOL;
	echo 'SheetName:' . $sheetName . EOL;
	echo 'current_days:' . $current_days . '天' . EOL . EOL;
	
	$objPHPExcel->setActiveSheetIndex( $sheetIndex );
	$activeSheet = $objPHPExcel->getActiveSheet();
	
	$activeSheet->setCellValue( 'B2', $year );
	if ( $sheetIndex >= 6 ) 
	{
		$activeSheet->setCellValue( 'B6', '交通费+早(6)+中(20)+晚(10)+零(80)' );
		$activeSheet->setCellValue( 'B7', '补余额' );
		$activeSheet->setCellValue( 'B16', '水费' );
		$activeSheet->setCellValue( 'B17', '电费' );
		$activeSheet->setCellValue( 'B18', '燃气费' );
		$activeSheet->setCellValue( 'B19', '互联网费' );
		$activeSheet->setCellValue( 'B20', '移动电话费' );
		$activeSheet->setCellValue( 'B21', '房费' );
		$activeSheet->setCellValue( 'B22', '交通费01' );
		$activeSheet->setCellValue( 'B23', '交通费02' );
		$activeSheet->setCellValue( 'B24', '理发' );
		
		$activeSheet->setCellValue( 'B25', '书籍费' );
		$activeSheet->setCellValue( 'B26', '' );
		$activeSheet->setCellValue( 'B27', '' );
		
		$activeSheet->setCellValue( 'B28', '服装费用' );
		$activeSheet->setCellValue( 'B29', '' );
		$activeSheet->setCellValue( 'B30', '' );
		$activeSheet->setCellValue( 'B31', '' );
		
		// 日期值显示格式
		$date_format = 'n-j';
		$activeSheet->setCellValue( 'D6', date( $date_format, $beginday_timestamp ) );
		$activeSheet->setCellValue( 'D7', date( $date_format, $finishday_timestamp ) );
		
		// 单元格字体水平居右
		$activeSheet->getStyle( 'D6:D7' )->applyFromArray(
										array(
											'alignment' => array(
											'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
										)
									)
								);
		
		
		// 根据不同月份生成所在月预算
		$traffic_budget = 100 * 2;
		$breakfast_budget = $current_days * 6;
		$lunch_budget = $current_days * 20;
		$supper_budget = $current_days * 10;
		$snack_budget = 80;

		$month_budget = $traffic_budget + $breakfast_budget + $lunch_budget + $supper_budget + $snack_budget;
		$activeSheet->setCellValue( 'C6', $month_budget );
	}
	
	$money_format = '￥#,##0.00';
	// 设置金额显示格式
	$activeSheet->getStyle( 'C6:C10' )->getNumberFormat()->setFormatCode( $money_format );
	$activeSheet->getStyle( 'C16:C37' )->getNumberFormat()->setFormatCode( $money_format );
	
	$activeSheet->getStyle( 'G6:G10' )->getNumberFormat()->setFormatCode( $money_format );
	$activeSheet->getStyle( 'F16' )->getNumberFormat()->setFormatCode( $money_format );
	$activeSheet->getStyle( 'F20' )->getNumberFormat()->setFormatCode( $money_format );
	
	// 设置日期格式
	$activeSheet->getStyle( 'D6:D9' )->getNumberFormat()->setFormatCode( 'm-d' );
	$activeSheet->getStyle( 'D16:D36' )->getNumberFormat()->setFormatCode( 'm-d' );
	
	$activeSheet->getColumnDimension( 'C' )->setWidth( 12 );
	$activeSheet->getColumnDimension( 'G' )->setWidth( 11 );
	
	$i = 1;
	for ( $column = 'J'; $column <= 'Z'; $column++ )
	{
		if ( 0 == $i % 2 )
		{
			//echo '字符:' . $column . ',Ascii值:' . ord( $column ) . EOL;
			
			$s_coordinate = $column . '4';
			$e_coordinate = $column . '42';
			
			$cellCoordinate = $s_coordinate . ':' . $e_coordinate;
			// echo 'cellCorrdinate:' . $cellCoordinate . EOL;
			
			$activeSheet->getStyle( $cellCoordinate )->getNumberFormat()->setFormatCode( $money_format );
			$activeSheet->getColumnDimension( $column )->setWidth( 11 );
		}
		
		if ( $current_days == $i / 2 )
		{
			
			$activeSheet->getStyle( 'F16' )->getFont()->getColor()->setARGB( PHPExcel_Style_Color::COLOR_RED );
			
			// echo '结束列:' . $column . EOL . EOL;
			$formula_express = "=SUM('1月:$sheetName'!F16:G17)";
			
			// echo '余额公式:' . $formula_express . EOL. EOL;
			
			$activeSheet->setCellValue( 'F20', $formula_express );
			$activeSheet->getStyle( 'F20' )->getAlignment()->setWrapText( true );
			$activeSheet->getStyle( 'F20' )->getFont()->getColor()->setARGB( PHPExcel_Style_Color::COLOR_RED );
			
			/* 
			 * 边框设置效果不好看...
			 * $styleThickGreenBorderOutline = array(
				'borders' => array( 'outline' => array( 'style' => PHPExcel_Style_Border::BORDER_THICK,
														'color' => array( 'argb' => PHPExcel_Style_Color::COLOR_GREEN ),
														),
									),
				);
			$activeSheet->getStyle( 'F20:G21' )->applyFromArray( $styleThickGreenBorderOutline ); 
			*/
			
			break;
		}
	
		$i ++;
	}
		
}

$objWriter = new PHPExcel_Writer_Excel5( $objPHPExcel );
$objWriter->setPreCalculateFormulas( false );
$objWriter->save( 'new.xls' );
