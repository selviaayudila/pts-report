<?php
// Dapatkan nama file saat ini
$current_page = basename($_SERVER['PHP_SELF']);
?>

<aside class="main-sidebar sidebar-light-danger elevation-4" style="height: 100%; position: fixed; top: 0; left: 0; width: 250px; overflow-y: auto; z-index: 1030;">
  <a href="dash.php" class="brand-link" style="display: flex; justify-content: center; align-items: center; height: 57px;">
    <img src="dist/img/logo-text.png" alt="" class="brand-image" style="opacity: .8;">
  </a>
  <div class="sidebar" style="overflow-x: hidden; white-space: nowrap;">
    <!-- Sidebar user panel -->
    <div class="user-panel mt-3 pb-3 mb-3 d-flex justify-content-center align-items-center text-center">
      <div class="info" style="white-space: normal; word-wrap: break-word;">
        <a href="#" class="d-block">Production Reporting System</a>
      </div>
    </div>

    <!-- Sidebar Search Form -->
    <div class="form-inline">
      <div class="input-group" data-widget="sidebar-search">
        <input class="form-control form-control-sidebar" type="search" placeholder="Search" aria-label="Search" style="width: 100%; box-sizing: border-box;">
        <div class="input-group-append">
          <button class="btn btn-sidebar" style="width: auto;">
            <i class="fas fa-search fa-fw"></i>
          </button>
        </div>
      </div>
    </div>

    <!-- Sidebar Menu -->
    <nav class="mt-2">
      <ul class="nav nav-pills nav-sidebar flex-column" data-widget="treeview" role="menu" data-accordion="false">
        <!-- Dashboard -->
        <li class="nav-item">
          <a href="index.php" class="nav-link <?= $current_page == 'index.php' ? 'active' : '' ?>" style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
            <i class="nav-icon fas fa-edit"></i>
            <p>Report</p>
          </a>
        </li>

        <!-- Users -->
      
      </ul>
    </nav>
  </div>
</aside>
