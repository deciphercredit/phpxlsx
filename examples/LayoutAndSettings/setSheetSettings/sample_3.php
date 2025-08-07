<?php
// modify the page layout of the Chart sheet name in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

$xlsx->setActiveSheet(array('name' => 'Chart'));
$xlsx->setSheetSettings(array('paperType' => 'A3'));

$xlsx->saveXlsx('example_setSheetSettings_3');