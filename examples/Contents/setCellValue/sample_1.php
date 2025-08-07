<?php
// set cell values in cell positions in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

// change the active sheet
$xlsx->setActiveSheet(array('position' => 1));

// set the cell value in B5 position
$content = array(
    'text' => 'New content',
);
$xlsx->setCellValue($content, 'B5');
// the same can be done using addCell, that allows using more options
// $xlsx->addCell($content, 'B5', array(), array('useCellStyles' => true));

$xlsx->saveXlsx('example_setCellValue_1');