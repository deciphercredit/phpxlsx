<?php
// add headers in sheets in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum',
);
$xlsx->addCell($content, 'A1');

$headerContent = array(
    'text' => 'First sheet header',
);

$xlsx->addHeader(array('center' => $headerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

// select the new sheet as the active one
$xlsx->addSheet(array('name'=> 'Other', 'removeSelected' => true, 'selected' => true, 'active' => true));

$headerContent = array(
    'text' => 'Other sheet header',
);

$xlsx->addHeader(array('center' => $headerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addHeader_5');