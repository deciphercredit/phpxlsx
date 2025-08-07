<?php
// add cell contents applying styles in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
);
$xlsx->addCell($content, 'A1');
$content = array(
    'text' => 'Lorem ipsum',
);
$xlsx->addCell($content, 'A2');

$content = array(
    'text' => 'Sed ut perspiciatis unde omnis',
    'italic' => true,
    'underline' => 'single',
);
$xlsx->addCell($content, 'B5');

$content = array(
    'text' => 'At vero eos et accusamus et iusto',
    'color' => '13775F',
);
$xlsx->addCell($content, 'F2');

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
    'bold' => true,
    'font' => 'Times New Roman',
    'strikethrough' => true,
);
$xlsx->addCell($content, 'AA3');

$xlsx->saveXlsx('example_addCell_1');