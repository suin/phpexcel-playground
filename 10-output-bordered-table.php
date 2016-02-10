<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

$book = new PHPExcel();
$sheet = $book->getActiveSheet();
$sheet
    ->getStyle('A1:C3')
    ->getBorders()
    ->getAllBorders()
    ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

// 色を指定する場合
$sheet
    ->getStyle('A5:C8')
    ->applyFromArray([
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => ['rgb' => 'FF0000'],
            ],
        ],
    ]);

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/10-罫線で表を描く.xlsx');
