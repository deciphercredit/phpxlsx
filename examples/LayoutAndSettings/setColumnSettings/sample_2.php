<?php
// set custom column width values in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

// set column settings. The position is set using column letters
$xlsx->setColumnSettings('A', array('heighwidtht' => 20));
$xlsx->setColumnSettings('B', array('width' => 20.25));
$xlsx->setColumnSettings('C', array('width' => 30));
$xlsx->setColumnSettings('D', array('width' => 20));
$xlsx->setColumnSettings('E', array('width' => 20));

// set the third sheet as the active sheet
$xlsx->setActiveSheet(array('position' => 2));

// set column settings. The position is set using cell positions
$xlsx->setColumnSettings('A1', array('width' => 30));

$xlsx->saveXlsx('example_setColumnSettings_2');