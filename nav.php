<?php
// Pastikan session_start() dipanggil sebelum ada output HTML
session_start();
?>

<style>
  /* Tambahkan CSS untuk membuat navbar tetap di atas */
  .main-header {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    z-index: 1030; /* Pastikan navbar berada di atas konten lainnya */
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  }

  /* Tambahkan margin-top pada konten untuk menghindari tumpang tindih dengan navbar */
  .content-wrapper {
    margin-top: 56px; /* Sesuaikan dengan tinggi navbar */
  }
</style>

<nav class="main-header navbar navbar-expand navbar-white navbar-light">
  <!-- Left navbar links -->
  <ul class="navbar-nav">
    <li class="nav-item">
      <a class="nav-link" data-widget="pushmenu" href="#" role="button"><i class="fas fa-bars"></i></a>
    </li>
    
    <li class="nav-item dropdown">
  
    
</li>

  </ul>

  <!-- Right navbar links -->
  <ul class="navbar-nav ml-auto">
    <!-- Navbar Search -->
   
    <!-- Fullscreen -->
    <li class="nav-item">
      <a class="nav-link" data-widget="fullscreen" href="#" role="button">
        <i class="fas fa-expand-arrows-alt"></i>
      </a>
    </li>

    <!-- User Profile and Logout -->
   
  </ul>
</nav>

