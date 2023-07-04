<?php

// Melde alle PHP-Fehler
error_reporting(-1);

// Include excel function
include "include_excel.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

// Load the existing Excel file
$spreadsheet = IOFactory::load("vorlagen/Mitglieder.xlsx");

// Get the active worksheet
$worksheet = $spreadsheet->getActiveSheet();

// Set the value of cell A1 to "Hello World!"
$worksheet->setCellValue("A1", "Hello World!");

// Save the updated file
$writer = IOFactory::createWriter($spreadsheet, "Xlsx");

// return Xlsx
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="example.xlsx"');
header('Cache-Control: max-age=0');
$writer->save('php://output');
// $writer->save("hello_world.xlsx");
