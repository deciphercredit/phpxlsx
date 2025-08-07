<?php
// add area charts in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

// contents to be used as legends
$xlsx->addCell(array('text' => 'legend 1'), 'A2');
$xlsx->addCell(array('text' => 'legend 2'), 'B2');
$xlsx->addCell(array('text' => 'legend 3'), 'C2');

// contents to be used as labels
$xlsx->addCell(array('text' => 'data 1'), 'A3');
$xlsx->addCell(array('text' => 'data 2'), 'B3');
$xlsx->addCell(array('text' => 'data 3'), 'C3');
$xlsx->addCell(array('text' => 'data 4'), 'D3');

// contents to be used as values
$xlsx->addCell(array('text' => 10), 'A4');
$xlsx->addCell(array('text' => 20), 'B4');
$xlsx->addCell(array('text' => 50), 'C4');
$xlsx->addCell(array('text' => 25), 'D4');
$xlsx->addCell(array('text' => 7), 'A5');
$xlsx->addCell(array('text' => 60), 'B5');
$xlsx->addCell(array('text' => 33), 'C5');
$xlsx->addCell(array('text' => 0), 'D5');
$xlsx->addCell(array('text' => 5), 'A6');
$xlsx->addCell(array('text' => 3), 'B6');
$xlsx->addCell(array('text' => 7), 'C6');
$xlsx->addCell(array('text' => 14), 'D6');

$content = array(
    'text' => 'A 2D area chart:',
    'bold' => true,
);
$xlsx->addCell($content, 'A9');

$chartData = array(
    'legends' => array('A2', 'B2', 'C2'),
    'labels' => array('A3', 'B3', 'C3', 'D3'),
    'values' => array(
        array('A4', 'B4', 'C4', 'D4'),
        array('A5', 'B5', 'C5', 'D5'),
        array('A6', 'B6', 'C6', 'D6'),
    ),
);

$optionsChart = array(
    'data' => $chartData,
    'showTable' => false,
    'legendPos' => 'b',
    'legendOverlay' => false,
    'hgrid' => 3,
    'vgrid' => 1,
    'colSize' => 10,
    'rowSize' => 18,
);
$xlsx->addChart('area', 'H7', $optionsChart);

$content = array(
    'text' => 'A 3D area chart:',
    'bold' => true,
);
$xlsx->addCell($content, 'A30');

$chartData = array(
    'legends' => array('A2', 'B2', 'C2'),
    'labels' => array('A3', 'B3', 'C3', 'D3'),
    'values' => array(
        array('A4', 'B4', 'C4', 'D4'),
        array('A5', 'B5', 'C5', 'D5'),
        array('A6', 'B6', 'C6', 'D6'),
    ),
);

$optionsChart = array(
    'data' => $chartData,
    'color' => '2',
    'perspective' => 30,
    'rotX' => 30,
    'rotY' => 30,
    'font' => 'Arial',
    'showTable' => true,
    'legendPos' => 'r',
    'legendOverlay' => false,
    'hgrid' => 3,
    'vgrid' => 2,
    'colSize' => 10,
    'rowSize' => 18,
);
$xlsx->addChart('area3D', 'H28', $optionsChart);

$xlsx->saveXlsx('example_addChart_6');