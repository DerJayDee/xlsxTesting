<?php
// Include the PhpSpreadsheet classes
require_once "SimpleCache/CacheInterface.php";
require_once "PhpOffice/PhpSpreadsheet/IComparable.php";
require_once "PhpOffice/PhpSpreadsheet/ReferenceHelper.php";
require_once "PhpOffice/PhpSpreadsheet/Collection/CellsFactory.php";
require_once "PhpOffice/PhpSpreadsheet/Collection/Cells.php";
require_once "PhpOffice/PhpSpreadsheet/Collection/Memory/SimpleCache1.php";
require_once "PhpOffice/PhpSpreadsheet/Calculation/Calculation.php";
require_once "PhpOffice/PhpSpreadsheet/Calculation/Category.php";
require_once "PhpOffice/PhpSpreadsheet/Calculation/Engine/BranchPruner.php";
require_once "PhpOffice/PhpSpreadsheet/Calculation/Engine/CyclicReferenceStack.php";
require_once "PhpOffice/PhpSpreadsheet/Calculation/Engine/Logger.php";
require_once "PhpOffice/PhpSpreadsheet/Settings.php";
require_once "PhpOffice/PhpSpreadsheet/Spreadsheet.php";
require_once "PhpOffice/PhpSpreadsheet/Reader/Security/XmlScanner.php";
require_once "PhpOffice/PhpSpreadsheet/Reader/IReadFilter.php";
require_once "PhpOffice/PhpSpreadsheet/Reader/DefaultReadFilter.php";
require_once "PhpOffice/PhpSpreadsheet/Reader/IReader.php";
require_once "PhpOffice/PhpSpreadsheet/Reader/BaseReader.php";
require_once "PhpOffice/PhpSpreadsheet/Reader/Xlsx/Namespaces.php";
require_once "PhpOffice/PhpSpreadsheet/Reader/Xlsx.php";
require_once "PhpOffice/PhpSpreadsheet/Worksheet/Worksheet.php";
require_once "PhpOffice/PhpSpreadsheet/Writer/IWriter.php";
require_once "PhpOffice/PhpSpreadsheet/Writer/BaseWriter.php";
require_once "PhpOffice/PhpSpreadsheet/Writer/Xlsx.php";
require_once "PhpOffice/PhpSpreadsheet/Shared/File.php";
require_once "PhpOffice/PhpSpreadsheet/Shared/StringHelper.php";
require_once "PhpOffice/PhpSpreadsheet/IOFactory.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

// Load the existing Excel file
$spreadsheet = IOFactory::load("test.xlsx");

// Get the active worksheet
$worksheet = $spreadsheet->getActiveSheet();

// Set the value of cell A1 to "Hello World!"
$worksheet->setCellValue("A1", "Hello World!");

// Save the updated file
$writer = new Xlsx($spreadsheet);
$writer->save("test.xlsx");

echo "Cell A1 updated successfully.";
?>
