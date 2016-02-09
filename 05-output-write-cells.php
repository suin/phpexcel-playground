<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

$book = new PHPExcel();
$sheet = $book->getActiveSheet();

// セル番地で書いてみる
$sheet->setCellValue('A1', 10000);
$sheet->setCellValue('A2', true);
$sheet->setCellValue('A3', 'テスト');

// 行列番号で書いてみる
$column = 1;
$sheet->setCellValueByColumnAndRow($column, 1, 'B1');
$sheet->setCellValueByColumnAndRow($column, 2, 'B2');
$sheet->setCellValueByColumnAndRow($column, 3, 'B3');

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/05-セルに書いてみる.xlsx');
