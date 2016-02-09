<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

$book = new PHPExcel();
$book->getActiveSheet()->setTitle("シート1です");
$book->createSheet()->setTitle("二枚目!");
$book->createSheet()->setTitle("さんまいめだよ");

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/02-シートに名前をつける.xlsx');
