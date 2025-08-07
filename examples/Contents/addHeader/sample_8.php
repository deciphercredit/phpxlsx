<?php
// add an image in the first page and an image and text contents as default header for other pages applying options in an XLSX created from scratch

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

// first header
$headerImageContent = array(
    'image' => array(
        'src' => '../../files/image.png',
        'height' => 30,
        'width' => 40,
        'title' => 'Image header',
        'alt' => 'Image header',
    ),
);
$xlsx->addHeader(array('center' => $headerImageContent), 'first');

// default header
$headerContentPage = array(
    'text' => '&[Page] of &[Pages]',
);
$headerContentDefault = array(
    'text' => 'Default header',
);
// image content
$headerImageContent = array(
    'image' => array(
        'src' => '../../files/image.png',
        'color' => 'washout',
    ),
);
$xlsx->addHeader(array('left' => $headerContentDefault, 'center' => $headerContentPage, 'right' => $headerImageContent));

// set view as pageLayout to display headers and footers
$xlsx->setSheetSettings(array('view' => 'pageLayout'));

$xlsx->saveXlsx('example_addHeader_8');