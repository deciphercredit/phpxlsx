<?php
// add a doughnut chart in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

// contents to be used as legends
$xlsx->addCell(array('text' => 'legend 1'), 'A3');
$xlsx->addCell(array('text' => 'legend 2'), 'B3');
$xlsx->addCell(array('text' => 'legend 3'), 'C3');
$xlsx->addCell(array('text' => 'legend 4'), 'D3');
$xlsx->addCell(array('text' => 'legend 5'), 'E3');

// contents to be used as values
$xlsx->addCell(array('text' => 20), 'A4');
$xlsx->addCell(array('text' => 20), 'B4');
$xlsx->addCell(array('text' => 50), 'C4');
$xlsx->addCell(array('text' => 25), 'D4');
$xlsx->addCell(array('text' => 5), 'E4');

$content = array(
    'text' => 'A doughnut chart:',
    'bold' => true,
);
$xlsx->addCell($content, 'A6');

$chartData = array(
    'legends' => array('A3', 'B3', 'C3', 'D3', 'E3'),
    'values' => array('A4', 'B4', 'C4', 'D4', 'E4'),
);

$optionsChart = array(
    'data' => $chartData,
);
$xlsx->addChart('doughnut', 'H7', $optionsChart);

$xlsx->saveXlsx('example_addChart_5');