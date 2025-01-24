<?php
$host = '192.168.1.99'; // Replace with MySQL server IP
$port = 3306;
$user = 'root'; // MySQL user
$password = 'Qwer1234'; // MySQL password
$dbname = 'pts_db'; // Database name

// Create a connection
$mysqli = new mysqli($host, $user, $password, $dbname, $port);

// Check the connection
if ($mysqli->connect_error) {
    die("Connection failed: " . $mysqli->connect_error);
} else {
    echo "Successfully connected to the database '$dbname' on server '$host'.";
}
?>
