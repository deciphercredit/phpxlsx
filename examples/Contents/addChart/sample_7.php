<?php
// add a sheet with legends and values and another sheet with charts in an XLSX created from scratch

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

// add a new sheet
$xlsx->addSheet(array('name'=> 'Charts', 'removeSelected' => true, 'selected' => true, 'active' => true));

// add a chart using the data from Sheet1
$chartData = array(
    'legends' => array('Sheet1!A3', 'Sheet1!B3', 'Sheet1!C3', 'Sheet1!D3'),
    'values' => array('Sheet1!A4', 'Sheet1!B4', 'Sheet1!C4', 'Sheet1!D4'),
);

$optionsChart = array(
    'data' => $chartData,
    'color' => 3,
    'sizeX' => 10,
    'sizeY' => 5,
);
$xlsx->addChart('pie', 'H10', $optionsChart);

$xlsx->saveXlsx('example_addChart_7');