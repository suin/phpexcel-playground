<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

$products = [
    ['Microsoft Excel 2016',      14000, 2],
    ['Microsoft Word 2016',       14800, 1],
    ['Microsoft PowerPoint 2016', 15000, 1],
];

$book = PHPExcel_IOFactory::load('templates/03-見積書テンプレート.xltx');
$sheet = $book->getActiveSheet();

$rowOffset = 3;
foreach ($products as $row => $product) {
    foreach ($product as $col => $value) {
        $sheet->setCellValueByColumnAndRow($col, $row + $rowOffset, $value);
    }
}

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/03-テンプレートを使って出力.xlsx');
