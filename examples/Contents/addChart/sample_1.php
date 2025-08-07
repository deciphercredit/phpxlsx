<?php
// add pie charts in an XLSX created from scratch

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
    'text' => 'A 2D pie chart:',
    'bold' => true,
);
$xlsx->addCell($content, 'A6');

$chartData = array(
    'legends' => array('A3', 'B3', 'C3', 'D3'),
    'values' => array('A4', 'B4', 'C4', 'D4'),
);

$optionsChart = array(
    'data' => $chartData,
);
$xlsx->addChart('pie', 'H7', $optionsChart);

$content = array(
    'text' => 'A 3D pie chart:',
    'bold' => true,
);
$xlsx->addCell($content, 'A30');

$chartData = array(
    'legends' => array('A3', 'B3', 'C3', 'D3'),
    'values' => array('A4', 'B4', 'C4', 'D4'),
);

$optionsChart = array(
    'data' => $chartData,
    'rotX' => 20,
    'rotY' => 20,
    'perspective' => 30,
    'color' => '3',
    'showPercent' => true,
);
$xlsx->addChart('pie3D', 'H28', $optionsChart);

$xlsx->saveXlsx('example_addChart_1');