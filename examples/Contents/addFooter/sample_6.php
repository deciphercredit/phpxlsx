<?php
// add footers in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates.xlsx');

// add footers into the first sheet
$xlsx->setActiveSheet(array('position' => 0));

$footerContent = array(
    'text' => 'Footer first sheet',
);

$xlsx->addFooter(array('center' => $footerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

// add footers into the second sheet, replace existing footers
$xlsx->setActiveSheet(array('position' => 1));

$footerContent = array(
    'text' => 'Footer second sheet',
);

$xlsx->addFooter(array('center' => $footerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

// add footers into the fourth sheet, do not replace existing footers
$xlsx->setActiveSheet(array('position' => 3));

$footerContent = array(
    'text' => 'Footer fourth sheet',
);

$xlsx->addFooter(array('center' => $footerContent), 'default', array('replace' => false));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addFooter_6');