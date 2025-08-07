<?php
// add headers in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates.xlsx');

// add headers into the first sheet
$xlsx->setActiveSheet(array('position' => 0));

$headerContent = array(
    'text' => 'Header first sheet',
);

$xlsx->addHeader(array('center' => $headerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

// add headers into the second sheet, replace existing headers
$xlsx->setActiveSheet(array('position' => 1));

$headerContent = array(
    'text' => 'Header second sheet',
);

$xlsx->addHeader(array('center' => $headerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

// add headers into the fourth sheet, do not replace existing headers
$xlsx->setActiveSheet(array('position' => 3));

$headerContent = array(
    'text' => 'Header fourth sheet',
);

$xlsx->addHeader(array('center' => $headerContent), 'default', array('replace' => false));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addHeader_6');