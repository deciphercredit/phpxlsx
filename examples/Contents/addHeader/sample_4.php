<?php
// add contents with styles in a header for all pages in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
);
$xlsx->addCell($content, 'A1');

$headerContent = array(
    array(
        'text' => 'Header',
        'bold' => true,
    ),
    array(
        'text' => ' content.',
        'italic' => true,
    ),
);

$xlsx->addHeader(array('center' => $headerContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addHeader_4');