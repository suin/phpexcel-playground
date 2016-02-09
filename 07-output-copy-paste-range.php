<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

/**
 * 行を完全コピーする
 *
 * http://blog.kotemaru.org/old/2012/04/06.html より
 * @param PHPExcel_Worksheet $sheet
 * @param int $srcRow
 * @param int $dstRow
 * @param int $height
 * @param int $width
 * @throws PHPExcel_Exception
 */
function copyRows(
    PHPExcel_Worksheet $sheet,
    $srcRow,
    $dstRow,
    $height,
    $width
) {
    for ($row = 0; $row < $height; $row++) {
        // セルの書式と値の複製
        for ($col = 0; $col < $width; $col++) {
            $cell = $sheet->getCellByColumnAndRow($col, $srcRow + $row);
            $style = $sheet->getStyleByColumnAndRow($col, $srcRow + $row);

            $dstCell = PHPExcel_Cell::stringFromColumnIndex($col) . (string)($dstRow + $row);
            $sheet->setCellValue($dstCell, $cell->getValue());
            $sheet->duplicateStyle($style, $dstCell);
        }

        // 行の高さ複製。
        $h = $sheet->getRowDimension($srcRow + $row)->getRowHeight();
        $sheet->getRowDimension($dstRow + $row)->setRowHeight($h);
    }

    // セル結合の複製
    // - $mergeCell="AB12:AC15" 複製範囲の物だけ行を加算して復元。
    // - $merge="AB16:AC19"
    foreach ($sheet->getMergeCells() as $mergeCell) {
        $mc = explode(":", $mergeCell);
        $col_s = preg_replace("/[0-9]*/", "", $mc[0]);
        $col_e = preg_replace("/[0-9]*/", "", $mc[1]);
        $row_s = ((int)preg_replace("/[A-Z]*/", "", $mc[0])) - $srcRow;
        $row_e = ((int)preg_replace("/[A-Z]*/", "", $mc[1])) - $srcRow;

        // 複製先の行範囲なら。
        if (0 <= $row_s && $row_s < $height) {
            $merge = $col_s . (string)($dstRow + $row_s) . ":" . $col_e . (string)($dstRow + $row_e);
            $sheet->mergeCells($merge);
        }
    }
}

$book = PHPExcel_IOFactory::load('templates/07-交通費精算書テンプレート.xlsx');
$sheet = $book->getActiveSheet();

copyRows($sheet, 2, 7, 5, 5);
copyRows($sheet, 2, 7 + 5, 5, 5);
copyRows($sheet, 2, 7 + 10, 5, 5);
copyRows($sheet, 2, 7 + 15, 5, 5);

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/07-交通費精算書.xlsx');
