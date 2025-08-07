<?php
// modify the page layout of second and third sheets in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

$xlsx->setActiveSheet(array('position' => 1));
$xlsx->setSheetSettings(array('paperType' => 'letter-landscape'));

$xlsx->setActiveSheet(array('position' => 2));
$xlsx->setSheetSettings(array('paperType' => 'letter-landscape'));

$xlsx->saveXlsx('example_setSheetSettings_2');