<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// Fungsi untuk membaca file Excel
function readExcelFile($filePath) {
    if (file_exists($filePath)) {
        $spreadsheet = IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
        return $sheet->toArray();
    }
    return null;
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
    if ($shift === '3rd') {
        $time = strtotime($timestamp);
        if (date('H', $time) < 7) { // Jika sebelum pukul 07:00
            $date = date('d/m/Y', strtotime('-1 day', $time)); // Kurangi satu hari
        }
    }
    return $date;
}

function calculateDowntime($machineChangesData) {
    $downtimeData = [];
    $lastStatus = [];
    $lastTimestamp = [];
    $lastReason = [];

    foreach ($machineChangesData as $index => $row) {
        if ($index == 0 || !isset($row[1], $row[2], $row[3], $row[5])) continue; // Lewati header & validasi

        $machineID = trim($row[1]); // Kolom B (Machine ID)
        $status = trim($row[2]);   // Kolom C (Status)
        $downtimeReason = trim($row[3]); // Kolom D (Downtime Reason)
        $timestamp = trim($row[5]);     // Kolom F (Timestamp)

        if (empty($timestamp)) continue; // Lewati jika timestamp kosong

        $shift = tentukanShift($timestamp);
        $date = formatShiftDate($timestamp, $shift);

        if (!isset($downtimeData[$machineID][$date][$shift])) {
            $downtimeData[$machineID][$date][$shift] = [
                'totalDowntime' => 0,
                'downtimeReasons' => []
            ];
        }

        if (!isset($lastStatus[$machineID][$shift])) {
            if ($status === '0') {
                $lastStatus[$machineID][$shift] = '0';
                $lastTimestamp[$machineID][$shift] = $timestamp;
                $lastReason[$machineID][$shift] = $downtimeReason;
            }
            continue;
        }

        if ($status === '0' && $lastStatus[$machineID][$shift] === '1') {
            if (!empty($lastTimestamp[$machineID][$shift])) {
                $duration = (strtotime($timestamp) - strtotime($lastTimestamp[$machineID][$shift])) / 3600; // Durasi dalam jam
                $downtimeData[$machineID][$date][$shift]['totalDowntime'] += $duration;

                if (!isset($downtimeData[$machineID][$date][$shift]['downtimeReasons'][$downtimeReason])) {
                    $downtimeData[$machineID][$date][$shift]['downtimeReasons'][$downtimeReason] = 0;
                }
                $downtimeData[$machineID][$date][$shift]['downtimeReasons'][$downtimeReason] += $duration;
            }

            $lastStatus[$machineID][$shift] = '0';
            $lastTimestamp[$machineID][$shift] = $timestamp;
            $lastReason[$machineID][$shift] = $downtimeReason;
        } elseif ($status === '1') {
            $lastStatus[$machineID][$shift] = '1';
            $lastTimestamp[$machineID][$shift] = $timestamp;
        }
    }

    return $downtimeData;
}


// Ambil bulan dan tahun dari query string
$month = isset($_GET['month']) ? $_GET['month'] : null;
$year = isset($_GET['year']) ? $_GET['year'] : null;

if ($month && $year) {
    $folderPath = __DIR__ . "/pts_db/$year/$month/";

    // Path file Excel
    $filePaths = [
        'output' => $folderPath . "output_entries.xlsx",
        'jobOrder' => $folderPath . "job_order_changes.xlsx",
        'reject' => $folderPath . "reject_entries.xlsx",
        'toolingChanges' => $folderPath . "tooling_changes.xlsx",
        'toolingTargets' => $folderPath . "tooling_targets.xlsx",
        'products' => $folderPath . "products.xlsx",
        'tonnage' => $folderPath . "tonnage.xlsx",
        'machineChanges' => $folderPath . "machine_changes.xlsx",
    ];

    // Membaca semua file Excel menggunakan fungsi readExcelFile
    $outputData = readExcelFile($filePaths['output']);
    $jobOrderData = readExcelFile($filePaths['jobOrder']);
    $rejectData = readExcelFile($filePaths['reject']);
    $toolingData = readExcelFile($filePaths['toolingChanges']);
    $toolingTargetsData = readExcelFile($filePaths['toolingTargets']);
    $productsData = readExcelFile($filePaths['products']);
    $tonnageData = readExcelFile($filePaths['tonnage']);
    $machineChangesData = readExcelFile($filePaths['machineChanges']);

    // Pastikan semua data telah dibaca dengan benar
    if ($outputData && $machineChangesData) {
        // Hitung downtime menggunakan data dari machine_changes.xlsx
        $downtimeData = calculateDowntime($machineChangesData);

        // Tambahkan downtime ke hasil output
        foreach ($outputData as $index => $row) {
            if ($index == 0) continue; // Lewati header
            $machineID = trim($row[1]); // Kolom B untuk Machine ID
            $timestamp = trim($row[5]); // Kolom F untuk Timestamp
            $shiftDate = formatShiftDate($timestamp, tentukanShift($timestamp));
            $shift = tentukanShift($timestamp);

            $row['totalDowntime'] = $downtimeData[$machineID][$shiftDate][$shift]['totalDowntime'] ?? 0;
            $row['downtimeReasons'] = $downtimeData[$machineID][$shiftDate][$shift]['downtimeReasons'] ?? [];
            $outputData[$index] = $row;
        }

        // Map tooling targets berdasarkan toolingID dan productID
        $toolingTargets = [];
        foreach ($toolingTargetsData as $index => $row) {
            if ($index == 0) continue; // Lewati header

            $toolingID = $row[0];  // Kolom A (Tooling ID)
            $productID = $row[1];  // Kolom B (Product ID)
            $targetCavityNum = $row[2]; // Kolom C (Target Cavity Number)
            $targetCycleTime = $row[3]; // Kolom D (Target Cycle Time)

            if ($toolingID) {
                $toolingTargets[$toolingID] = [
                    'targetCavityNum' => $targetCavityNum,
                    'targetCycleTime' => $targetCycleTime
                ];
            }
            if ($productID) {
                $toolingTargets[$productID] = [
                    'targetCavityNum' => $targetCavityNum,
                    'targetCycleTime' => $targetCycleTime
                ];
            }
        }
// Map job order details
$jobOrderDetails = [];
foreach ($jobOrderData as $index => $row) {
    if ($index == 0) continue; // Lewati header

    $jobOrderID = $row[1]; // Kolom B (Job Order ID)
    $machineID = $row[2];  // Kolom C (Machine ID)
    $toolingID = $row[3];  // Kolom D (Tooling ID)
    $status = $row[4];     // Kolom E (Status)
    $timestamp = $row[6];  // Kolom G (Timestamp)

    // Format ulang tanggal ke d/m/y
    $formattedTimestamp = date('d/m/Y H:i', strtotime($timestamp));

    // Jika jobOrderID belum ada, inisialisasi
    if (!isset($jobOrderDetails[$jobOrderID])) {
        $jobOrderDetails[$jobOrderID] = [
            'machineID' => $machineID,
            'toolingID' => $toolingID,
            'startDate' => $formattedTimestamp, // Tetapkan timestamp pertama sebagai startDate
            'endDate' => 'RUNNING',            // Default endDate adalah RUNNING
        ];
    }

    // Perbarui endDate jika status adalah 'PAUSED' atau 'CLOSED'
    if (in_array($status, ['PAUSED', 'CLOSED'])) {
        $jobOrderDetails[$jobOrderID]['endDate'] = $formattedTimestamp;
    }
}

        // Map tooling changes to get productID, cavityNum, and cycleTime by toolingID
        $toolingDetails = [];
        foreach ($toolingData as $index => $row) {
            if ($index == 0) continue; // Lewati header

            $toolingID = $row[1];  // Kolom B (Tooling ID)
            $productID = $row[2];  // Kolom C (Product ID)
            $cavityNum = $row[3];  // Kolom D (Cavity Number)
            $cycleTime = $row[4];  // Kolom E (Cycle Time)

            $toolingDetails[$toolingID] = [
                'productID' => $productID,
                'cavityNum' => $cavityNum,
                'cycleTime' => $cycleTime,
            ];
        }

        // Map products untuk mendapatkan customerID dan productName berdasarkan productID
        $productDetails = [];
        foreach ($productsData as $index => $row) {
            if ($index == 0) continue; // Lewati header

            $productID = $row[0];  // Kolom A (Product ID)
            $customerID = $row[1]; // Kolom B (Customer ID)
            $productName = $row[2]; // Kolom C (Product Name)

            $productDetails[$productID] = [
                'customerID' => $customerID,
                'productName' => $productName,
            ];
        }

        // Map tonnage untuk mendapatkan tonnage berdasarkan machineID
        $machineTonnage = [];
        foreach ($tonnageData as $index => $row) {
            if ($index == 0) continue; // Lewati header

            $machineID = $row[0];  // Kolom A (Machine ID)
            $tonnage = $row[1];    // Kolom B (Tonnage)

            $machineTonnage[$machineID] = $tonnage;
        }

        // Proses output_entries.xlsx
        $results = [];
        foreach ($outputData as $index => $row) {
            if ($index == 0) continue; // Lewati header

            $jobOrderID = $row[2]; // Kolom C (Job Order ID)
            $quantity = $row[3];   // Kolom D (Quantity)
            $timestamp = $row[5];  // Kolom F (Timestamp)

            $shift = tentukanShift($timestamp);
            $shiftDate = formatShiftDate($timestamp, $shift);

            // Susun kunci berdasarkan jobOrderID, tanggal, dan shift
            $key = "$jobOrderID|$shiftDate|$shift";

            if (!isset($results[$key])) {
                $toolingID = $jobOrderDetails[$jobOrderID]['toolingID'] ?? 'N/A';
                $productID = $toolingDetails[$toolingID]['productID'] ?? 'N/A';
                $machineID = $jobOrderDetails[$jobOrderID]['machineID'] ?? 'N/A';
                $results[$key] = [
                    'date' => $shiftDate, // PostingDate
                    'customerID' => $productDetails[$productID]['customerID'] ?? 'N/A', // CustName
                    'productID' => $productID, // PartNo
                    'productName' => $productDetails[$productID]['productName'] ?? 'N/A', // PartName
                    'toolingID' => $toolingID, // MoldNo
                    'jobOrderID' => $jobOrderID, // JSNo
                    'shift' => $shift, // OprShift
                    'totalQuantity' => 0, // AOutput
                    'cavityNum' => $toolingDetails[$toolingID]['cavityNum'] ?? 'N/A', // Acavity
                    'targetCavityNum' => $toolingTargets[$toolingID]['targetCavityNum'] ?? $toolingTargets[$productID]['targetCavityNum'] ?? 'N/A', // QCV
                    'targetCycleTime' => $toolingTargets[$toolingID]['targetCycleTime'] ?? $toolingTargets[$productID]['targetCycleTime'] ?? 'N/A', // QCT
                    'cycleTime' => $toolingDetails[$toolingID]['cycleTime'] ?? 'N/A', // ACT
                    'ShiftHour' => 7.5, // ShiftHour (default value)
                    'NumOpr' => 1, // NumOpr (default value)
                    'machineID' => $machineID, // ActualMC
                    'tonnage' => $machineTonnage[$machineID] ?? 'N/A', // MCTonnage
                    'startDate' => $jobOrderDetails[$jobOrderID]['startDate'] ?? 'N/A', // StartTime
                    'endDate' => $jobOrderDetails[$jobOrderID]['endDate'] ?? 'N/A', // EndTime
                    'rejects' => [], // Untuk jenis reject
                    
                    'downtimeReasons' => [], // Rincian downtime
                ];
            }

            $results[$key]['totalQuantity'] += $quantity;
        }

        // Proses reject_entries.xlsx
        foreach ($rejectData as $index => $row) {
            if ($index == 0) continue; // Lewati header

            $jobOrderID = $row[1]; // Kolom B (Job Order ID)
            $rejectType = $row[2]; // Kolom C (Reject Type)
            $quantity = $row[3];   // Kolom D (Quantity)
            $timestamp = $row[5];  // Kolom F (Timestamp)

            $shift = tentukanShift($timestamp);
            $shiftDate = formatShiftDate($timestamp, $shift);

            $key = "$jobOrderID|$shiftDate|$shift";

            if (isset($results[$key])) {
                if (!isset($results[$key]['rejects'][$rejectType])) {
                    $results[$key]['rejects'][$rejectType] = 0;
                }
                $results[$key]['rejects'][$rejectType] += $quantity;
            }
        }
        
        foreach ($downtimeData as $machineID => $shiftData) {
            foreach ($shiftData as $shiftDate => $shifts) {
                foreach ($shifts as $shift => $downtime) {
                    foreach ($results as $key => &$result) {
                        // Pastikan data konsisten
                        $result['machineID'] = trim($result['machineID']);
                        $result['date'] = trim($result['date']);
                        $result['shift'] = trim($result['shift']);
        
                        $machineID = trim($machineID);
                        $shiftDate = trim($shiftDate);
                        $shift = trim($shift);
        
                        // Cocokkan data
                        if (
                            $result['machineID'] === $machineID &&
                            $result['date'] === $shiftDate &&
                            $result['shift'] === $shift
                        ) {
                            // Hindari akumulasi berulang
                            if (!isset($result['processed']) || !$result['processed']) {
                                foreach ($downtime['downtimeReasons'] as $reason => $duration) {
                                    // Penanganan khusus untuk JOS dan EOS
                                    if ($reason === 'JOS' || $reason === 'EOS') {
                                        // Biarkan JOS dan EOS dihitung penuh
                                        if (!isset($result['downtimeReasons'][$reason])) {
                                            $result['downtimeReasons'][$reason] = 0;
                                        }
                                        $result['downtimeReasons'][$reason] = number_format(
                                            $result['downtimeReasons'][$reason] + $duration,
                                            1
                                        );
                                    } else {
                                        // Penanganan normal untuk alasan lainnya
                                        if (!isset($result['downtimeReasons'][$reason])) {
                                            $result['downtimeReasons'][$reason] = 0;
                                        }
                                        $result['downtimeReasons'][$reason] = number_format(
                                            $result['downtimeReasons'][$reason] + min($duration, 7.5), 
                                            1
                                        );
                                    }
                                }
        
                                $result['processed'] = true;
                            }
                        }
                    }
                }
            }
        }
        

        
    } else {
        echo "File output_entries.xlsx, job_order_changes.xlsx, reject_entries.xlsx, tooling_changes.xlsx, tooling_targets.xlsx, atau downtime_entries.xlsx tidak ditemukan di folder $folderPath.";
    }
} else {
    echo "Silakan pilih bulan dan tahun terlebih dahulu.";
}
if (!$machineChangesData) {
    echo "Data machine_changes.xlsx tidak dapat dibaca atau kosong.";
}

?>
