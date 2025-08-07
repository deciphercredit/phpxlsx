<?php
// add first, default and even headers in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
);
$xlsx->addCell($content, 'A1');
$content = array(
    'text' => 'J1 content',
);
$xlsx->addCell($content, 'J1');
$content = array(
    'text' => 'S1 content',
);
$xlsx->addCell($content, 'S1');

// empty content
$headerContentFirst = array(
    'text' => '',
);
// first header
$xlsx->addHeader(array('center' => $headerContentFirst), 'first');

// add special elements and plain text contents
$headerContentDateTime = array(
    'text' => 'Date and time: &[Date] &[Time]',
);
$headerContentPage = array(
    'text' => '&[Page] of &[Pages]',
);
$headerContentDefault = array(
    'text' => 'Default header',
);

// default header
$xlsx->addHeader(array('left' => $headerContentDateTime, 'center' => $headerContentDefault, 'right' => $headerContentPage));

// even header
$xlsx->addHeader(array('left' => $headerContentDateTime, 'right' => $headerContentPage), 'even');

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addHeader_3');