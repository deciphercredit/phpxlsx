<?php
// add an image in the first page and an image and text contents as default footer for other pages applying options in an XLSX created from scratch

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

// first footer
$footerImageContent = array(
    'image' => array(
        'src' => '../../files/image.png',
        'height' => 30,
        'width' => 40,
        'title' => 'Image footer',
        'alt' => 'Image footer',
    ),
);
$xlsx->addFooter(array('center' => $footerImageContent), 'first');

// default footer
$footerContentPage = array(
    'text' => '&[Page] of &[Pages]',
);
$footerContentDefault = array(
    'text' => 'Default footer',
);
// image content
$footerImageContent = array(
    'image' => array(
        'src' => '../../files/image.png',
        'color' => 'washout',
    ),
);
$xlsx->addFooter(array('left' => $footerContentDefault, 'center' => $footerContentPage, 'right' => $footerImageContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addFooter_8');