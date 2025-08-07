<?php
// replace chart values

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Utilities\XlsxUtilities();

$data = array(
    array(
        'row' => 2,
        'col' => 1,
        'value' => 30,
    ),
    array(
        'row' => 2,
        'col' => 2,
        'value' => 10,
    ),
    array(
        'row' => 2,
        'col' => 3,
        'value' => 5,
    ),
    array(
        'row' => 4,
        'col' => 2,
        'value' => 3,
    ),
);

$xlsx->searchAndReplace('../../files/data_excel.xlsx', 'example_searchAndReplace_2.xlsx', $data, 'sheet', array('sheetName' => 'Chart'));
//$xlsx->searchAndReplace('../../files/data_excel.xlsx', 'example_searchAndReplace_2.xlsx', $data, 'sheet', array('sheetNumber' => 3));