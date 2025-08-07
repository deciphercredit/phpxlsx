<?php
// add left, center and right contents applying styles in a header for all pages in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
);
$xlsx->addCell($content, 'A1');

$headerContentLeft = array(
    'text' => 'My header',
    'color' => 'FF0000',
    'italic' => true,
    'font' => 'Times New Roman',
    'fontSize' => 24,
);

$headerContentCenter = array(
    'text' => 'Header content',
    'bold' => true,
);

$headerContentRight = array(
    'text' => 'Another header',
    'strikethrough' => true,
    'underline' => 'single',
);

$xlsx->addHeader(array('left' => $headerContentLeft, 'center' => $headerContentCenter, 'right' => $headerContentRight));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addHeader_2');