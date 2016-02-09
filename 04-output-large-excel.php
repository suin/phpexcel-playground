<?php
date_default_timezone_set('Asia/Tokyo');
ini_set('memory_limit', '1024M');

require __DIR__ . '/vendor/autoload.php';
$book = new PHPExcel();
$sheet = $book->getActiveSheet();
$startedOn = time();

for ($row = 1; $row <= 100000; $row++) {
    $sheet->setCellValueByColumnAndRow(0, $row, str_repeat('あ', 32));
    if ($row % 10000 === 0) {
        $writer2007 = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
        $writer2007->save('output/04-large-excel.xlsx');
        $writer5 = PHPExcel_IOFactory::createWriter($book, 'Excel5');
        $writer5->save('output/04-large-excel.xls');
        rap($row, $startedOn);
    }
}

function rap($row, $startedOn)
{
    printf(
        "% 4d sec| rows: % 6d xlsx: % 4uKB xls: % 4uKB mem: %.02fMB\n",
        time() - $startedOn,
        $row,
        filesize('output/04-large-excel.xlsx') / 1024,
        filesize('output/04-large-excel.xls') / 1024,
        memory_get_usage(true) / 1024 / 1024
    );
}