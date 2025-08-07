<?php
// set custom row height values in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

// set row settings. The position is set using row numbers
$xlsx->setRowSettings(1, array('height' => 60));
$xlsx->setRowSettings(2, array('height' => 50.25));
$xlsx->setRowSettings(3, array('height' => 70));
$xlsx->setRowSettings(4, array('height' => 50));
$xlsx->setRowSettings(5, array('height' => 50));

// set the third sheet as the active sheet
$xlsx->setActiveSheet(array('position' => 2));

// set row settings. The position is set using cell positions
$xlsx->setRowSettings('A2', array('height' => 10));
$xlsx->setRowSettings('A1', array('height' => 10));

$xlsx->saveXlsx('example_setRowSettings_2');