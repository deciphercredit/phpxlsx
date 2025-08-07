<?php
// add a table in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

// add new contents
$contents = array(
    array(
        10,
    ),
    array(
        20,
    ),
    array(
        55,
    ),
    array(
        32,
    ),
);
$options = array(
    'columnNames' => array(
        array(
            'text' => 'Values',
            'bold' => true,
        ),
    ),
);
$xlsx->addTable($contents, 'B3', array(), $options);

$xlsx->saveXlsx('example_addTable_1');