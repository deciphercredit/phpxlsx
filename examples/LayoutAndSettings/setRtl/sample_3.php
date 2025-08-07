<?php
// enable rtl of second and third sheet in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

$xlsx->setActiveSheet(array('position' => 1));
$xlsx->setSheetSettings(array('rtl' => true));

$xlsx->setActiveSheet(array('position' => 2));
$xlsx->setSheetSettings(array('rtl' => true));

$xlsx->saveXlsx('example_setRtl_3');