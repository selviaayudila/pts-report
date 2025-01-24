<?php
require 'vendor/autoload.php'; // Pastikan PHPSpreadsheet terpasang
require 'config.php';  // Sertakan file koneksi database Anda

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Array Tonnage
$tonnageData = [
    "A1" => 170, "A2" => 450, "A3" => 350, "A5" => 350, "A6" => 350, "A7" => 230,
    "A8" => 280, "A9" => 280, "A10" => 280, "A11" => 550, "A12" => 550, "A13" => 350,
    "A14" => 350, "A15" => 350, "B1" => 230, "B2" => 220, "B3" => 200, "B4" => 200,
    "B5" => 230, "B6" => 230, "B7" => 120, "B8" => 120, "B9" => 550, "B10" => 220,
    "B11" => 150, "B12" => 230, "B13" => 60, "B16" => 450, "B17" => 220, "C1" => 180,
    "C2" => 180, "C3" => 180, "C4" => 180, "C5" => 180, "C6" => 80, "C7" => 180,
    "C8" => 180, "C9" => 180, "C10" => 180, "C11" => 180, "C12" => 180, "C13" => 100,
    "C14" => 100, "C15" => 100, "C16" => 75, "C17" => 100, "C18" => 100, "C19" => 130,
    "D4" => 40, "D5" => 110, "D6" => 75, "D7" => 75, "D8" => 60, "D9" => 40,
    "E1" => 50, "E2" => 80, "E3" => 80, "E5" => 50, "E6" => 75, "E7" => 75,
    "E8" => 30, "E9" => 40, "E10" => 80, "E11" => 80, "E12" => 60, "E13" => 80,
    "E15" => 40, "E16" => 35, "F1" => 280, "F2" => 280, "F3" => 130, "F5" => 130,
    "F6" => 130, "G1" => 80, "G2" => 80, "G3" => 80, "G10" => 35, "G5" => 80,
    "G6" => 80, "G7" => 80, "G8" => 50, "G9" => 35, "G11" => 60, "G12" => 50,
    "D3" => 60, "B13" => 50, "D9" => 80,
];

// Validasi input 'month' dan 'year'
$month = isset($_POST['month']) ? $_POST['month'] : null;
$year = isset($_POST['year']) ? $_POST['year'] : null;

// Jika 'month' atau 'year' tidak ada, hentikan proses
if (!$month || !$year) {
    die(" Please select month now! ");
}

// Tentukan rentang waktu untuk tabel job_order_changes (satu bulan sebelum hingga akhir bulan yang diminta)
$previousMonth = $month == 1 ? 12 : $month - 1; // Bulan sebelumnya
$previousYear = $month == 1 ? $year - 1 : $year; // Tahun jika bulan Januari
$startDateJobOrder = "$previousYear-" . str_pad($previousMonth, 2, '0', STR_PAD_LEFT) . "-01 07:00:00";
$nextMonth = $month == 12 ? 1 : $month + 1; // Bulan berikutnya
$nextYear = $month == 12 ? $year + 1 : $year; // Tahun jika bulan Desember
$endDateJobOrder = "$nextYear-" . str_pad($nextMonth, 2, '0', STR_PAD_LEFT) . "-01 07:00:00";

// Tentukan rentang waktu untuk tabel lainnya (satu bulan penuh)
$startDate = "$year-$month-01 07:00:00";
$endDate = "$nextYear-" . str_pad($nextMonth, 2, '0', STR_PAD_LEFT) . "-01 07:00:00";

// Tentukan folder untuk menyimpan file
$baseDir = __DIR__ . "/pts_db/$year/$month/";
if (!file_exists($baseDir)) {
    mkdir($baseDir, 0777, true); // Buat folder jika belum ada
}

// **Pembuatan File Tonnage**
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('tonnage');

// Header kolom
$sheet->setCellValue('A1', 'machineID');
$sheet->setCellValue('B1', 'tonnage');

// Isi data
$rowIndex = 2;
foreach ($tonnageData as $machineID => $tonnage) {
    $sheet->setCellValue("A$rowIndex", $machineID);
    $sheet->setCellValue("B$rowIndex", $tonnage . 'T'); // Tambahkan "T" di belakang nilai tonase
    $rowIndex++;
}


// Simpan file tonnage
$filePathTonnage = $baseDir . "tonnage.xlsx";
$writerTonnage = new Xlsx($spreadsheet);
$writerTonnage->save($filePathTonnage);

// Proses tabel job_order_changes
$query = "SELECT * FROM job_order_changes WHERE timestamp BETWEEN '$startDateJobOrder' AND '$endDateJobOrder'";
$result = $mysqli->query($query); // Gunakan koneksi dari file config.php

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('job_order_changes');

// Tambahkan header dan data
if ($result && $result->num_rows > 0) {
    $columns = array_keys($result->fetch_assoc()); // Ambil nama kolom
    $result->data_seek(0); // Kembalikan pointer hasil ke baris pertama
    $colIndex = 1;
    foreach ($columns as $column) {
        $sheet->setCellValueByColumnAndRow($colIndex++, 1, $column);
    }

    $rowIndex = 2;
    while ($row = $result->fetch_assoc()) {
        $colIndex = 1;
        foreach ($row as $value) {
            $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex, $value !== null && $value !== '' ? $value : 'NULL');
        }
        $rowIndex++;
    }
}

$filePath = $baseDir . "job_order_changes.xlsx";
$writer = new Xlsx($spreadsheet);
$writer->save($filePath);

// Proses tabel lainnya
$filteredTables = ['machine_changes', 'output_entries', 'reject_entries', 'tooling_changes'];
foreach ($filteredTables as $tableName) {
    // Tambahkan ORDER BY dan LIMIT hanya untuk tooling_changes
    $additionalClause = $tableName === 'tooling_changes' ? "ORDER BY timestamp LIMIT 1000000" : "";

    $query = "SELECT * FROM $tableName WHERE timestamp BETWEEN '$startDate' AND '$endDate' $additionalClause";
    $result = $mysqli->query($query);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle($tableName);

    if ($result && $result->num_rows > 0) {
        $columns = array_keys($result->fetch_assoc()); // Ambil nama kolom
        $result->data_seek(0);
        $colIndex = 1;
        foreach ($columns as $column) {
            $sheet->setCellValueByColumnAndRow($colIndex++, 1, $column);
        }

        $rowIndex = 2;
        while ($row = $result->fetch_assoc()) {
            $colIndex = 1;
            foreach ($row as $value) {
                $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex, $value !== null && $value !== '' ? $value : 'NULL');
            }
            $rowIndex++;
        }
    }

    $filePath = $baseDir . "$tableName.xlsx";
    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);
}

// Proses tabel tanpa filter rentang waktu
$noFilterTables = ['products', 'tooling_targets'];
foreach ($noFilterTables as $tableName) {
    $query = "SELECT * FROM $tableName";
    $result = $mysqli->query($query);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle($tableName);

    if ($result && $result->num_rows > 0) {
        $columns = array_keys($result->fetch_assoc());
        $result->data_seek(0);
        $colIndex = 1;
        foreach ($columns as $column) {
            $sheet->setCellValueByColumnAndRow($colIndex++, 1, $column);
        }

        $rowIndex = 2;
        while ($row = $result->fetch_assoc()) {
            $colIndex = 1;
            foreach ($row as $value) {
                $sheet->setCellValueByColumnAndRow($colIndex++, $rowIndex, $value !== null && $value !== '' ? $value : 'NULL');
            }
            $rowIndex++;
        }
    }

    $filePath = $baseDir . "$tableName.xlsx";
    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);
}

// Berikan pesan sukses
echo " File successfully created in folder: $baseDir ";
?>
