<?php
// set tabSelected in sheets and activeCell in the workbook options in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');
$xlsx->setSheetSettings(array('tabSelected' => false, 'activeCell' => 'B2'));

$xlsx->setActiveSheet(array('position' => 2));
$xlsx->setSheetSettings(array('tabSelected' => true, 'activeCell' => 'D4'));

// set active sheet as the active tab in the workbook
$xlsx->setWorkbookSettings(array('activeSheetAsActiveTab' => true));

$xlsx->saveXlsx('example_setSheetSettings_4');