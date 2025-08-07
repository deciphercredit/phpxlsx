<?php
// add first, default and even footers in an XLSX created from scratch

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
$footerContentFirst = array(
    'text' => '',
);
// first footer
$xlsx->addFooter(array('center' => $footerContentFirst), 'first');

// add special elements and plain text contents
$footerContentDateTime = array(
    'text' => 'Date and time: &[Date] &[Time]',
);
$footerContentPage = array(
    'text' => '&[Page] of &[Pages]',
);
$footerContentDefault = array(
    'text' => 'Default footer',
);

// default footer
$xlsx->addFooter(array('left' => $footerContentDateTime, 'center' => $footerContentDefault, 'right' => $footerContentPage));

// even footer
$xlsx->addFooter(array('left' => $footerContentDateTime, 'right' => $footerContentPage), 'even');

// set view as pageLayout to display footers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addFooter_3');