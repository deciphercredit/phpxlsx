<?php
// add left, center and right contents applying styles in a footer for all pages in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
);
$xlsx->addCell($content, 'A1');

$footerContentLeft = array(
    'text' => 'My footer',
    'color' => 'FF0000',
    'italic' => true,
    'font' => 'Times New Roman',
    'fontSize' => 24,
);

$footerContentCenter = array(
    'text' => 'Footer content',
    'bold' => true,
);

$footerContentRight = array(
    'text' => 'Another footer',
    'strikethrough' => true,
    'underline' => 'single',
);

$xlsx->addFooter(array('left' => $footerContentLeft, 'center' => $footerContentCenter, 'right' => $footerContentRight));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addFooter_2');