<?php
// add a table with columns setting multiple row totals in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

// contents to be added
$contents = array(
    array(
        10,
        50,
    ),
    array(
        20,
        25.5,
    ),
    array(
        55,
        12,
    ),
    array(
        32,
        36,
    ),
);
$tableStyles = array(
    'totalRow' => true,
);

// table with max and count as row totals
$options = array(
    'columnNames' => array(
        array(
            'text' => 'Values 1',
        ),
        array(
            'text' => 'Values 2',
        ),
    ),
    'columnTotals' => array(
        array(
            'type' => 'function',
            'value' => 'max',
            'contentStyles' => array(
                'bold' => true,
            ),
        ),
        array(
            'type' => 'function',
            'value' => 'count',
            'contentStyles' => array(
                'bold' => true,
            ),
        ),
    ),
);
$xlsx->addTable($contents, 'B3', $tableStyles, $options);

$xlsx->saveXlsx('example_addTable_4');