<?php
// add a pie chart with a title styled in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

// contents to be used as legends
$xlsx->addCell(array('text' => 'legend 1'), 'A3');
$xlsx->addCell(array('text' => 'legend 2'), 'B3');
$xlsx->addCell(array('text' => 'legend 3'), 'C3');
$xlsx->addCell(array('text' => 'legend 4'), 'D3');

// contents to be used as values
$xlsx->addCell(array('text' => 10), 'A4');
$xlsx->addCell(array('text' => 20), 'B4');
$xlsx->addCell(array('text' => 50), 'C4');
$xlsx->addCell(array('text' => 25), 'D4');

$content = array(
    'text' => 'A 3D pie chart with a title styled:',
    'bold' => true,
);
$xlsx->addCell($content, 'A6');

$chartData = array(
    'legends' => array('A3', 'B3', 'C3', 'D3'),
    'values' => array('A4', 'B4', 'C4', 'D4'),
);

$optionsChart = array(
    'data' => $chartData,
    'title' => 'My title',
    'perspective' => 30,
    'color' => '2',
    'showPercent' => true,
    'vgrid' => 0,
    'legendPos' => 'r',
    'font' => 'Arial',
    'stylesTitle' => array(
        'bold' => true,
        'color' => 'FF0000',
        'font' => 'Times New Roman',
        'fontSize' => 3600,
        'italic' => true,
    ),
);
$xlsx->addChart('pie', 'H7', $optionsChart);

$xlsx->saveXlsx('example_addChart_8');