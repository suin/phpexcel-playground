<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

$book = new PHPExcel();
$sheet = $book->getActiveSheet();
$sheet->setCellValue('A1', 64);
$sheet->setCellValue('B1', 4);
$sheet->setCellValue('C1', '=A1 * B1');

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/12-計算式を書き込む.xlsx');
