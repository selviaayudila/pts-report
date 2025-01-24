<?php include('job.php'); ?>

<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="icon" href="dist/img/logo.png" type="image/png">
  <title>Sanwa | Report</title>

  <!-- Google Font: Source Sans Pro -->
  <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Source+Sans+Pro:300,400,400i,700&display=fallback">
  <!-- Font Awesome -->
  <link rel="stylesheet" href="plugins/fontawesome-free/css/all.min.css">
  <!-- DataTables -->
  <link rel="stylesheet" href="plugins/datatables-bs4/css/dataTables.bootstrap4.min.css">
  <link rel="stylesheet" href="plugins/datatables-responsive/css/responsive.bootstrap4.min.css">
  <link rel="stylesheet" href="plugins/datatables-buttons/css/buttons.bootstrap4.min.css">
  <!-- Theme style -->
  <link rel="stylesheet" href="dist/css/adminlte.min.css">
</head>
<body>

<?php include('nav.php'); ?>
<?php include('sidebar.php'); ?>

<div class="content-wrapper">
  <section class="content-header">
    <div class="container-fluid">
      <div class="row mb-2">
        <div class="col-sm-6">
          <h1>Monthly Report by JobOrderID</h1>
        </div>
        <div class="col-sm-6">
          <ol class="breadcrumb float-sm-right">
            <li class="breadcrumb-item"><a href="#">Report</a></li>
            <li class="breadcrumb-item active">Report View</li>
          </ol>
        </div>
      </div>
    </div>
  </section>

  <section class="content">
    <div class="card">
      <div class="card-body">
        <table id="jobOrderTable" class="table table-bordered table-striped">
          <thead>
            <tr>
              <th>PostingDate</th>
              <th>CustName</th>
              <th>PartNo</th>
              <th>PartName</th>
              <th>MoldNo</th>
              <th>JSNo</th>
              <th>OprShift</th>
              <th>AOutput</th>
              <th>ACavity</th>
              <th>QCV</th>
              <th>QCT</th>
              <th>ACT</th>
              <th>ShiftHours</th>
              <th>NumOpr</th>
              <th>ActualMC</th>
              <th>MCTonnage</th>
              <th>Start Time</th>
              <th>End Time</th>
              <th>Silver Streak</th>
              <th>Short Molding</th>
              <th>Dented</th>
              <th>Sink Mark</th>
              <th>Burn Mark</th>
              <th>Bubble</th>
              <th>Black Dot</th>
              <th>Scratches</th>
              <th>Flow Mark</th>
              <th>Dim Out</th>
              <th>Discolouration</th>
              <th>Shiny</th>
              <th>White Mark</th>
              <th>Flashes</th>
              <th>Drag Mark</th>
              <th>Oily Mark</th>
              <th>Over Cut</th>
              <th>Pin Mark</th>
              <th>Wrinkle</th>
              <th>Weld Line</th>
              <th>Pin Broken</th>
              <th>Damage</th>
              <th>Metal Chip</th>
              <th>Crack</th>
              <th>Part Drop to Floor</th>
              <th>Total Reject</th>
              <th>NPO</th>
              <th>TPM</th>
              <th>MOC</th>
              <th>T</th>
              <th>MSSD</th>
              <th>DMC</th>
              <th>NMP</th>
              <th>PA</th>
              <th>QP</th>
              <th>MB</th>
              <th>PM</th>
              <th>MOR1</th>
              <th>MOR2</th>
              <th>JOS</th>
              <th>EOS</th>
              
              
            </tr>
          </thead>
          <tbody>
            <?php if (!empty($results)): ?>
              <?php foreach ($results as $data): ?>
                <tr>
                  <td><?= htmlspecialchars($data['date'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['customerID'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['productID'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['productName'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['toolingID'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['jobOrderID'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['shift'] ?? '-') ?></td>
                  <td><?= htmlspecialchars(number_format($data['totalQuantity'], 0, '.', '.')) ?></td>
                  <td><?= htmlspecialchars($data['cavityNum'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['targetCavityNum'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['targetCycleTime'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['cycleTime'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['ShiftHour'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['NumOpr'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['machineID'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['tonnage'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['startDate'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['endDate'] ?? '-') ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Silver Streak'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Short Molding'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Dented'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Sink Mark'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Burn Mark'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Buble'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Black Dot'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Scratches'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Flow Mark'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Dim Out'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Discolouration'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Shiny'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['White M'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Flahes'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Drag Mark'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Oily Mark'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Over Cut'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Pin Mark'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Wrinkle'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Weld Line'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Pin broken'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Damage'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Metal Chip'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Crack'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['rejects']['Part Drop to Floor'] ?? 0) ?></td>
                  <td><?= htmlspecialchars(array_sum($data['rejects'] ?? [])) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['NPO'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['TPM'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['MOC'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['T'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['MS/SD'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['DMC'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['NMP'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['PA'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['QP'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['MB'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['PM'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['MOR1'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['MOR2'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['JOS'] ?? 0) ?></td>
                  <td><?= htmlspecialchars($data['downtimeReasons']['EOS'] ?? 0) ?></td>
              
                </tr>
              <?php endforeach; ?>
            <?php else: ?>
              <tr>
                <td colspan="43" style="text-align: center;">No data available</td>
              </tr>
            <?php endif; ?>
          </tbody>
        </table>
      </div>
    </div>
  </section>
</div>

<!-- Footer -->
<?php include 'footer.php'; ?>

<!-- Scripts -->
<script src="plugins/jquery/jquery.min.js"></script>
<script src="plugins/bootstrap/js/bootstrap.bundle.min.js"></script>
<script src="plugins/datatables/jquery.dataTables.min.js"></script>
<script src="plugins/datatables-bs4/js/dataTables.bootstrap4.min.js"></script>
<script src="plugins/datatables-responsive/js/dataTables.responsive.min.js"></script>
<script src="plugins/datatables-buttons/js/dataTables.buttons.min.js"></script>
<script src="plugins/datatables-buttons/js/buttons.bootstrap4.min.js"></script>
<script src="plugins/jszip/jszip.min.js"></script>
<script src="plugins/pdfmake/pdfmake.min.js"></script>
<script src="plugins/pdfmake/vfs_fonts.js"></script>
<script src="plugins/datatables-buttons/js/buttons.html5.min.js"></script>
<script src="plugins/datatables-buttons/js/buttons.print.min.js"></script>
<script src="plugins/datatables-buttons/js/buttons.colVis.min.js"></script>
<script src="dist/js/adminlte.min.js"></script>

<script>
   $(document).ready(function() {
    $('#jobOrderTable').DataTable({
        scrollX: true,
        autoWidth: false,
        lengthChange: true,
        searching: true,
        paging: true,
        info: true,
        buttons: [
            {
                extend: "excelHtml5",
                text: "Excel",
                exportOptions: {
                    columns: ':visible', // Ekspor hanya kolom yang terlihat
                    format: {
                        body: function(data, row, column, node) {
                            // Hapus format titik atau koma untuk kolom "Total Quantity"
                            return column === 7 ? data.replace(/\./g, '').replace(/,/g, '') : data;
                        }
                    }
                }
            },
            "csv", "copy", "print"
        ]
    }).buttons().container().appendTo('#jobOrderTable_wrapper .col-md-6:eq(0)');
});

$('#dataTable').css({
      'table-layout': 'fixed',
      'word-wrap': 'break-word'
    });

    $('#filterButton').click(function () {
      const form = $('#filterForm');
      if (form[0].checkValidity()) {
        // Serialize form data
        const formData = form.serializeArray();
        
        // Get month and year values from the form
        const month = formData.find(item => item.name === 'month').value;
        const year = formData.find(item => item.name === 'year').value;

        // Show a loading message or spinner
        const loadingMessage = $('<p>Loading, please wait...</p>').appendTo('.modal-body');

        // Use AJAX to send the form data
        $.post('process.php', $.param(formData), function (response) {
          // Remove loading message
          loadingMessage.remove();
          // Handle the response
          alert(`Database data (${month} ${year}) successfully exported!`);
          location.reload(); // Reload to reflect changes if needed
        }).fail(function (xhr, status, error) {
          // Remove loading message
          loadingMessage.remove();
          alert('An error occurred: ' + xhr.status + ' ' + error);
        });
      } else {
        alert('Please select both month and year!');
      }
    });
  

  document.getElementById("filterReportButton").addEventListener("click", function() {
    document.getElementById("selectMonthYearForm").submit();
});

</script>
</body>
</html>