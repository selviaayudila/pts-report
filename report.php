<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// File paths
$jobOrderFile = 'C:\xampp\htdocs\workbench\pts_db\2024\12/job_order_changes.xlsx';
$toolingChangeFile = 'C:\xampp\htdocs\workbench\pts_db\2024\12/tooling_changes.xlsx';
$productFile = 'C:\xampp\htdocs\workbench\pts_db\2024\12/products.xlsx';
$toolingTargetFile = 'C:\xampp\htdocs\workbench\pts_db\2024\12/tooling_targets.xlsx';
$tonnageFile = 'C:\xampp\htdocs\workbench\pts_db\2024\12/tonnage.xlsx';
$outputEntriesFile = 'C:\xampp\htdocs\workbench\pts_db\2024\12/output_entries.xlsx';
$rejectEntriesFile = 'C:\xampp\htdocs\workbench\pts_db\2024\12/reject_entries.xlsx'; 
$machineChangesFile = 'C:\xampp\htdocs\workbench\pts_db\2024\12/machine_changes.xlsx';
// Function to read Excel file
function readExcelFile($filePath) {
    $spreadsheet = IOFactory::load($filePath);
    $sheet = $spreadsheet->getActiveSheet();
    $data = $sheet->toArray(null, true, true, true);
    array_shift($data);  // Remove header
    return $data;
}
// Reading data from files
$jobOrderData = readExcelFile($jobOrderFile);
$toolingChangeData = readExcelFile($toolingChangeFile);
$productData = readExcelFile($productFile);
$toolingTargetData = readExcelFile($toolingTargetFile);
$tonnageData = readExcelFile($tonnageFile);
$outputEntriesData = readExcelFile($outputEntriesFile);
$rejectEntriesData = readExcelFile($rejectEntriesFile); // Reading reject entries
$machineChangesData = readExcelFile($machineChangesFile);

// Fungsi untuk mengurutkan data berdasarkan kolom timestamp
function sortDataByTimestamp($data, $timestampColumn) {
    usort($data, function($a, $b) use ($timestampColumn) {
        $timeA = strtotime($a[$timestampColumn]);
        $timeB = strtotime($b[$timestampColumn]);
        return $timeA <=> $timeB;
    });
    return $data;
}
// Urutkan machineChangesData berdasarkan timestamp sebelum menghitung downtime
$machineChangesData = sortDataByTimestamp($machineChangesData, 'F');

// Fungsi utama untuk menghitung downtime
function calculateDowntime($machineChangesData) {
    $downtimeData = [];
    $lastStatus = [];
    $lastTimestamp = [];
    $lastReason = [];
    $setA = [];
    $setB = [];

    if (empty($machineChangesData)) {
        return $downtimeData;
    }

    foreach ($machineChangesData as $row) {
        $machineID = trim($row['B']);
        $status = trim($row['C']);
        $downtimeReason = trim($row['D']);
        $timestamp = trim($row['F']);
        $shift = tentukanShift($timestamp);
        $date = date('d/m/Y', strtotime($timestamp));

        if (!isset($downtimeData[$machineID][$date][$shift])) {
            $downtimeData[$machineID][$date][$shift] = [
                'totalDowntime' => 0,
                'downtimeReasons' => []
            ];
        }

        // Kondisi #1: Tidak ada data sebelumnya
        if (!isset($lastStatus[$machineID][$shift])) {
            if ($status === '0') {
                $lastStatus[$machineID][$shift] = '0';
                $lastTimestamp[$machineID][$shift] = $timestamp;
                $lastReason[$machineID][$shift] = $downtimeReason;
            }
            continue;
        }

        // Kondisi #2: Baris tunggal data
        if ($status === '1' && $lastStatus[$machineID][$shift] === '0') {
            // Hitung durasi downtime untuk kondisi #2
            $lastTime = strtotime($lastTimestamp[$machineID][$shift]);
            $currentTime = strtotime($timestamp);
            $downtimeDuration = round(($currentTime - $lastTime) / 3600, 1);

            if (!isset($downtimeData[$machineID][$date][$shift]['downtimeReasons'][$lastReason[$machineID][$shift]])) {
                $downtimeData[$machineID][$date][$shift]['downtimeReasons'][$lastReason[$machineID][$shift]] = 0;
            }
            $downtimeData[$machineID][$date][$shift]['downtimeReasons'][$lastReason[$machineID][$shift]] += $downtimeDuration;
            $downtimeData[$machineID][$date][$shift]['totalDowntime'] += $downtimeDuration;

            $lastStatus[$machineID][$shift] = '1';
        } elseif ($status === '0' && $lastStatus[$machineID][$shift] === '1') {
            // Kondisi #3: Transisi ke downtime, tambahkan ke Set A
            $setA[] = [
                'timestamp' => $timestamp,
                'reason' => $downtimeReason
            ];
            $lastTimestamp[$machineID][$shift] = $timestamp;
            $lastReason[$machineID][$shift] = $downtimeReason;
            $lastStatus[$machineID][$shift] = '0';
        } elseif ($status === '0') {
            // Masukan downtime tambahan ke Set B untuk kondisi #3
            $setB[] = [
                'timestamp' => $timestamp,
                'reason' => $downtimeReason
            ];
        }
    }

    // Hitung downtime berdasarkan transisi di Set A dan Set B setelah loop terakhir

    return $downtimeData;
}


$downtimeData = calculateDowntime($machineChangesData);

// Fungsi untuk mengonversi array ke array asosiatif
function convertToAssocArray($data, $keyColumn) {
    $assocArray = [];
    foreach ($data as $row) {
        $key = trim($row[$keyColumn]);
        if (!isset($assocArray[$key])) {
            $assocArray[$key] = $row;
        }
    }
    return $assocArray;
}


$toolingChanges = convertToAssocArray($toolingChangeData, 'B');
$products = convertToAssocArray($productData, 'A');
$toolingTargets = convertToAssocArray($toolingTargetData, 'A');
$machineTonnage = convertToAssocArray($tonnageData, 'A');

// Use formatShiftDate to determine the correct date without splitting shift 3 across two days
foreach ($outputEntriesData as $row) {
    $jobOrderID = trim($row['C']); // Job Order ID
    $quantity = trim($row['D']); // Quantity/Output
    $timestamp = trim($row['F']); // Timestamp
    $shift = tentukanShift($timestamp);
    $date = formatShiftDate($timestamp, $shift); // Use new function to determine correct date

    // Use a combination of job order ID, date, and shift as the key
    $key = $jobOrderID . '|' . $date . '|' . $shift;

    if (!isset($outputEntries[$key])) {
        $outputEntries[$key] = 0; // Initialize if not set
    }
    $outputEntries[$key] += (int)$quantity; // Sum quantities
}


// Define specific reject types
$specificRejectTypes = [
    'Silver Streak', 'Short Molding', 'Dented', 'Sink Mark', 'Burn Mark', 'Buble', 
    'Black Dot', 'Scratches', 'Flow Mark', 'Dim Out', 'Discolouration', 'Shiny', 
    'White M', 'Flahes', 'Drag Mark', 'Oily Mark', 'Over Cut', 'Pin Mark', 
    'Wrinkle', 'Weld Line', 'Pin broken', 'Damage', 'Metal Chip', 'Others'
];

// Process reject entries by job order, date, and shift
$rejectEntries = [];
$dynamicRejectTypes = []; // To store different reject types found in the file

foreach ($rejectEntriesData as $row) {
    $jobOrderID = trim($row['B']); // Job Order ID
    $rejectType = trim($row['C']); // Reject Type
    $rejectQuantity = (int)trim($row['D']); // Convert Reject Quantity to integer
    $timestamp = trim($row['F']); // Timestamp
    $shift = tentukanShift($timestamp);
    $date = formatShiftDate($timestamp, $shift); // Adjusted for 3rd shift

    // Add reject type to dynamicRejectTypes array if it's not already present
    if (!in_array($rejectType, $dynamicRejectTypes) && !in_array($rejectType, $specificRejectTypes)) {
        $dynamicRejectTypes[] = $rejectType;
    }

    // Use a combination of job order ID, date, and shift as the key
    $key = $jobOrderID . '|' . $date . '|' . $shift;

    // Initialize the reject array for the key if not already set
    if (!isset($rejectEntries[$key])) {
        $rejectEntries[$key] = [];

        // Initialize all reject types with 0
        foreach (array_merge($specificRejectTypes, $dynamicRejectTypes) as $type) {
            $rejectEntries[$key][$type] = 0;
        }
    }

    // Add the reject quantity to the corresponding reject type for the key
    if (!isset($rejectEntries[$key][$rejectType])) {
        $rejectEntries[$key][$rejectType] = 0;
    }

    $rejectEntries[$key][$rejectType] += $rejectQuantity;
}

// Fungsi untuk menentukan shift berdasarkan timestamp
function tentukanShift($timestamp) {
    $time = strtotime($timestamp);
    $jam = date('H', $time);

    if ($jam >= 7 && $jam < 15) {
        return '1st';
    } elseif ($jam >= 15 && $jam < 23) {
        return '2nd';
    } else {
        return '3rd';
    }
}

// Fungsi untuk memformat tanggal shift dengan penyesuaian shift 3
function formatShiftDate($timestamp, $shift) {
    $date = date('d/m/Y', strtotime($timestamp));

    // Jika shift ke-3, pertahankan tanggal awal (tidak berpindah ke hari berikutnya)
    if ($shift === '3rd') {
        $time = strtotime($timestamp);

        // Jika timestamp melewati tengah malam tapi sebelum dini hari, sesuaikan tanggal ke 1 hari sebelumnya
        if (date('H', $time) < 7) {
            $date = date('d/m/Y', strtotime('-1 day', $time));
        }
    }

    return $date;
}

// Fungsi untuk mengurutkan data berdasarkan tanggal
function sortDataByDate($data, $dateColumn) {
    usort($data, function($a, $b) use ($dateColumn) {
        $dateA = strtotime($a[$dateColumn]);
        $dateB = strtotime($b[$dateColumn]);
        return $dateA <=> $dateB;
    });
    return $data;
}

// Urutkan data berdasarkan tanggal di kolom G
$jobOrderData = sortDataByDate($jobOrderData, 'G');

// Array untuk melacak end dates
$endDates = [];

// Populate end dates untuk job orders dengan status paused/closed
foreach ($jobOrderData as $jobOrder) {
    $status = strtolower(trim($jobOrder['E'])); // Status column
    $jobOrderID = trim($jobOrder['B']);
    $timestamp = strtotime($jobOrder['G']); // Start time

    if ($status == 'closed' || $status == 'paused') {
        $formattedEndDate = date('d/m/Y H:i', $timestamp);
        $endDates[$jobOrderID] = $formattedEndDate; // Simpan end date menggunakan job order ID
    }
}

// Fungsi untuk mengekspor data ke Excel
function exportToExcel($jobOrderData, $toolingChanges, $products, $toolingTargets, $machineTonnage, $endDates, $outputEntries, $rejectEntries, $specificRejectTypes, $dynamicRejectTypes, $downtimeData) {
    // Buat Spreadsheet baru
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Inisialisasi header tabel
    $headers = [
        'Posting Date', 'Customer Name', 'Part No', 'Part Name', 'Mold No', 'JS No', 'Opr Shift',
        'ACavity', 'Target Cavity (QCV)', 'Target Cycle Time (QCT)', 'Actual Cycle Time (ACT)', 
        'Shift Hour', 'Num Opr', 'Actual MC', 'MC Tonnage', 'Start Time', 'End Time', 'AOutput'
    ];

    // Tambahkan kolom reject types ke header
    foreach ($specificRejectTypes as $type) {
        $headers[] = $type;
    }
    foreach ($dynamicRejectTypes as $type) {
        $headers[] = $type;
    }
    $headers[] = 'Total Reject';

    // Tambahkan kolom downtime reasons ke header
    $uniqueDowntimeReasons = [];
    foreach ($downtimeData as $machineID => $dates) {
        foreach ($dates as $date => $shifts) {
            foreach ($shifts as $shift => $data) {
                foreach (array_keys($data['downtimeReasons']) as $reason) {
                    if (!in_array($reason, $uniqueDowntimeReasons)) {
                        $uniqueDowntimeReasons[] = $reason;
                    }
                }
            }
        }
    }
    foreach ($uniqueDowntimeReasons as $reason) {
        $headers[] = $reason;
    }
    $headers[] = 'Total Downtime';

    // Tulis header ke baris pertama
    $sheet->fromArray($headers, null, 'A1');

    // Inisialisasi indeks baris
    $rowIndex = 2;

    // Tulis data ke dalam Excel
    foreach ($outputEntries as $key => $outputQuantity) {
        list($jobOrderID, $date, $shift) = explode('|', $key);

        foreach ($jobOrderData as $jobOrder) {
            if (trim($jobOrder['B']) === $jobOrderID && strtolower(trim($jobOrder['E'])) === 'running') {
                // Isi data baris
                $postingDate = $date;
                $machineID = trim($jobOrder['C']);
                $toolingID = trim($jobOrder['D']);
                $timestamp = strtotime($jobOrder['G']);

                $productID = $toolingChanges[$toolingID]['C'] ?? 'N/A';
                $cavityNum = $toolingChanges[$toolingID]['D'] ?? 'N/A';
                $cycleTime = $toolingChanges[$toolingID]['E'] ?? 'N/A';

                $customerName = $products[$productID]['B'] ?? 'N/A';
                $productName = $products[$productID]['C'] ?? 'N/A';
                $targetCavity = $toolingTargets[$toolingID]['C'] ?? 'N/A';
                $targetCycleTime = $toolingTargets[$toolingID]['D'] ?? 'N/A';

                $startDate = date('d/m/Y H:i', $timestamp);
                $endDate = $endDates[$jobOrderID] ?? 'Running';
                $mcTonnage = ($machineTonnage[$machineID]['B'] ?? 'N/A') . ' T';

                // Data baris
                $rowData = [
                    $postingDate, $customerName, $productID, $productName, $toolingID, $jobOrderID, $shift,
                    $cavityNum, $targetCavity, $targetCycleTime, $cycleTime, 7.5, 1, $machineID, $mcTonnage,
                    $startDate, $endDate, $outputQuantity
                ];

                // Tambahkan reject data
                $rowRejectTotal = 0;
                foreach ($specificRejectTypes as $type) {
                    $rejectQuantity = $rejectEntries[$key][$type] ?? 0;
                    $rowData[] = $rejectQuantity;
                    $rowRejectTotal += $rejectQuantity;
                }
                foreach ($dynamicRejectTypes as $type) {
                    $rejectQuantity = $rejectEntries[$key][$type] ?? 0;
                    $rowData[] = $rejectQuantity;
                    $rowRejectTotal += $rejectQuantity;
                }
                $rowData[] = $rowRejectTotal;

                // Tambahkan downtime data
                $rowDowntimeTotal = 0;
                foreach ($uniqueDowntimeReasons as $reason) {
                    $downtimeDuration = $downtimeData[$machineID][$date][$shift]['downtimeReasons'][$reason] ?? 0;
                    $rowData[] = $downtimeDuration;
                    $rowDowntimeTotal += $downtimeDuration;
                }
                $rowData[] = $rowDowntimeTotal;

                // Tulis baris ke spreadsheet
                $sheet->fromArray($rowData, null, 'A' . $rowIndex);

                // Increment row index
                $rowIndex++;

                break;
            }
        }
    }



    // Simpan file Excel
    $writer = new Xlsx($spreadsheet);
    $filename = 'Monthly-Report.xlsx';

    // Atur header untuk download
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    header('Cache-Control: max-age=0');
    $writer->save('php://output');
    exit;
}
exportToExcel($jobOrderData, $toolingChanges, $products, $toolingTargets, $machineTonnage, $endDates, $outputEntries, $rejectEntries, $specificRejectTypes, $dynamicRejectTypes, $downtimeData);
?>