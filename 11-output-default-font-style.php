<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

$book = new PHPExcel();
$book
    ->getDefaultStyle()
    ->getFont()
    ->setName('メイリオ')
    ->setSize(16)
    ->setColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_GREEN));

$book->getActiveSheet()->setCellValue('A1', 'シート1');
$book->createSheet()->setCellValue('A1', 'シート2');

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/11-デフォルトのスタイル.xlsx');
