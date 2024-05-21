<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Color;

// สร้างวัตถุ Spreadsheet ใหม่
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// ตั้งค่าข้อมูลในเซลล์
$sheet->setCellValue('A1', 'Hello World!');

// ตั้งค่าสีตัวอักษรสำหรับเซลล์ A1
$sheet->getStyle('A1')->getFont()->getColor()->setARGB(Color::COLOR_RED);

// บันทึกไฟล์ Excel
$writer = new Xlsx($spreadsheet);
$writer->save('hello_world.xlsx');

echo "สร้างไฟล์ Excel เรียบร้อยแล้ว!";
?>
