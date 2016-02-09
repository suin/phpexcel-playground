<?php
date_default_timezone_set('Asia/Tokyo');

require __DIR__ . '/vendor/autoload.php';

$book = new PHPExcel();
$book->getProperties()
    ->setCreator("田中 太郎")
    ->setLastModifiedBy("山田 花子")
    ->setCompany('株式会社○○')
    ->setCreated(strtotime('2016-01-02 03:04:05'))
    ->setModified(strtotime('2016-02-03 04:05:06'))
    ->setManager('佐藤 次郎')
    ->setTitle("タイトル")
    ->setSubject("サブジェクト")
    ->setDescription("説明文")
    ->setKeywords("エクセル PHP 出力")
    ->setCategory("PHPすごい");

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/01-からっぽのエクセル.xlsx');
