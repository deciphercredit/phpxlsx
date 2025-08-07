<?php
// add contents with styles in a footer for all pages in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
);
$xlsx->addCell($content, 'A1');

$footerContent = array(
    array(
        'text' => 'Footer',
        'bold' => true,
    ),
    array(
        'text' => ' content',
        'italic' => true,
    ),
);

$xlsx->addFooter(array('center' => $footerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addFooter_4');