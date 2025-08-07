<?php
// add footers in sheets in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum',
);
$xlsx->addCell($content, 'A1');

$footerContent = array(
    'text' => 'First sheet footer',
);

$xlsx->addFooter(array('center' => $footerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

// select the new sheet as the active one
$xlsx->addSheet(array('name'=> 'Other', 'removeSelected' => true, 'selected' => true, 'active' => true));

$footerContent = array(
    'text' => 'Other sheet footer',
);

$xlsx->addFooter(array('center' => $footerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addFooter_5');