<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';
$book = new PHPExcel();
$sheet = $book->getActiveSheet();
$sheet->setCellValue('A1', 'テスト');

// xls: Excel97~2003
$writer2003 = PHPExcel_IOFactory::createWriter($book, 'Excel5');
$writer2003->save('output/08-excel2003.xls');

// xlsx: Excel2007~
$writer2007 = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer2007->save('output/08-excel2007.xlsx');
