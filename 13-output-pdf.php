<?php
date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

$products = [
    ['Microsoft Excel 2016',      14000, 2],
    ['Microsoft Word 2016',       14800, 1],
    ['Microsoft PowerPoint 2016', 15000, 1],
];

$book = PHPExcel_IOFactory::load('templates/13-見積書テンプレート.xlsx');
$sheet = $book->getActiveSheet();

$rowOffset = 3;
foreach ($products as $row => $product) {
    foreach ($product as $col => $value) {
        $sheet->setCellValueByColumnAndRow($col, $row + $rowOffset, $value);
    }
}

$excelWriter = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$excelWriter->save('output/13-excel.xlsx');

// tcPDF
PHPExcel_Settings::setPdfRenderer(
    PHPExcel_Settings::PDF_RENDERER_TCPDF,
    __DIR__ .'/vendor/tecnickcom/tcpdf'
);
$pdfWriter = PHPExcel_IOFactory::createWriter($book, 'PDF');
$pdfWriter->save('output/13-tcPDF.pdf');

// DomPDF
PHPExcel_Settings::setPdfRenderer(
    PHPExcel_Settings::PDF_RENDERER_DOMPDF,
    __DIR__ .'/vendor/dompdf/dompdf'
);
$pdfWriter = PHPExcel_IOFactory::createWriter($book, 'PDF');
$pdfWriter->save('output/13-Dompdf.pdf');

// mPDF
PHPExcel_Settings::setPdfRenderer(
    PHPExcel_Settings::PDF_RENDERER_MPDF,
    __DIR__ .'/vendor/mpdf/mpdf'
);
$pdfWriter = PHPExcel_IOFactory::createWriter($book, 'PDF');
$pdfWriter->save('output/13-mPDF.pdf');

// LibreOfficeでPDF化する
$soffice = '/Applications/LibreOffice.app/Contents/MacOS/soffice';
$outdir = __DIR__ . '/output';
$command = "$soffice --headless --convert-to pdf --outdir $outdir $outdir/13-excel.xlsx";
echo $command, PHP_EOL;
passthru($command);
