<?php
// add a CSV file setting the first row as header

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$options = array(
    'firstRowAsHeader' => true,
    'headerTextStyles' => array(
        'bold' => true,
        'italic' => true,
    ),
    'headerCellStyles' => array(
        'backgroundColor' => '4472C4',
    ),
);

$xlsx->addCsv('../../files/sample.csv', 'A1', $options);

$xlsx->saveXlsx('example_addCsv_2');