<?php
// hide sheets in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

// hide active sheet
$xlsx->setActiveSheet(array('position' => 0));
$xlsx->setSheetSettings(array('state' => 'hidden'));

// hide active sheet
$xlsx->setActiveSheet(array('position' => 2));
$xlsx->setSheetSettings(array('state' => 'hidden'));

// display active sheet
$xlsx->setSheetSettings(array('state' => 'visible'));

// the active tab can't be hidden. This value can be changed if needed
//$xlsx->setWorkbookSettings(array('activeTab' => 1));

$xlsx->saveXlsx('example_setSheetSettings_5');