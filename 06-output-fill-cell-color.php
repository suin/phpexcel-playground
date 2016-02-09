<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

$image = imagecreatefrompng('images/mc.png');
$width = imagesx($image);
$height = imagesy($image);

$book = new PHPExcel();
$sheet = $book->getActiveSheet();

for ($x = 0; $x < $width; $x ++) {
    $sheet->getColumnDimensionByColumn($x)->setWidth(1);
    for ($y = 0; $y < $height; $y ++) {
        $color = vsprintf('%02x%02x%02x', imagecolorsforindex($image, imagecolorat($image, $x, $y)));
        $sheet->getRowDimension($y + 1)->setRowHeight(6);
        $sheet
            ->getStyleByColumnAndRow($x, $y + 1)
            ->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()
            ->setRGB($color);
    }
}

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/06-セルの背景色.xlsx');
