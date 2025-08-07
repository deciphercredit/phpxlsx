<?php
// add a line chart with trendlines in an XLSX created from scratch

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
$xlsx->addCell(array('text' => 7), 'B4');
$xlsx->addCell(array('text' => 5), 'C4');
$xlsx->addCell(array('text' => 20), 'A5');
$xlsx->addCell(array('text' => 60), 'B5');
$xlsx->addCell(array('text' => 3), 'C5');
$xlsx->addCell(array('text' => 50), 'A6');
$xlsx->addCell(array('text' => 33), 'B6');
$xlsx->addCell(array('text' => 7), 'C6');
$xlsx->addCell(array('text' => 14), 'D6');

$content = array(
    'text' => 'A line chart with trendlines:',
    'bold' => true,
);
$xlsx->addCell($content, 'A9');

$chartData = array(
    'legends' => array('A2', 'B2', 'C2'),
    'labels' => array('A3', 'B3', 'C3', 'D3'),
    'values' => array(
        array('A4', 'B4', 'C4'),
        array('A5', 'B5', 'C5'),
        array('A6', 'B6', 'C6', 'D6'),
    ),
    'trendline' => array(
        array(
            'color' => '0000FF',
            'type' => 'log',
            'displayEquation' => true,
            'displayRSquared' => true,
        ),
        array(),
        array(
            'color' => '0000FF',
            'type' => 'power',
            'lineStyle' => 'dot',
        ),
    ),
);

$optionsChart = array(
    'data' => $chartData,
);

$xlsx->addChart('line', 'H7', $optionsChart);

$xlsx->saveXlsx('example_addChart_9');