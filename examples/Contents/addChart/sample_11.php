<?php
// add charts using cell ranges in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

// contents to be used as legends
$legends = array(
    array('legend 1', 'legend 2', 'legend 3'),
);
$xlsx->addCellRange($legends, 'A2');

// contents to be used as labels
$labels = array(
    array('data 1', 'data 2', 'data 3'),
);
$xlsx->addCellRange($labels, 'A3');

// contents to be used as values
$values = array(
    array(10, 20, 50),
    array(20, 6, 33),
    array(5, 4, 7),
);
$xlsx->addCellRange($values, 'A4');

// add a pie chart
$content = array(
    'text' => 'A pie chart:',
    'bold' => true,
);
$xlsx->addCell($content, 'A10');

$chartData = array(
    'legends' => array('A2:C2'),
    'values' => array('A4:C4'),
);

$optionsChart = array(
    'data' => $chartData,
);
$xlsx->addChart('pie', 'H7', $optionsChart);

// add a column chart
$content = array(
    'text' => 'A column chart:',
    'bold' => true,
);
$xlsx->addCell($content, 'A24');
$chartData = array(
    'legends' => array('A2:C2'),
    'labels' => array('A3:C3'),
    'values' => array(
        array('A4:C4'),
        array('A5:C5'),
        array('A6:C6'),
    ),
);

$optionsChart = array(
    'data' => $chartData,
    'color' => '3',
    'legendPos' => 'b',
);
$xlsx->addChart('col', 'H24', $optionsChart);

$xlsx->saveXlsx('example_addChart_11');