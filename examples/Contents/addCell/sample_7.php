<?php
// add cell contents in cell positions keeping existing styles in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

// change the active sheet
$xlsx->setActiveSheet(array('position' => 1));

// add a new content in B5 position setting the useCellStyles option as true
$content = array(
    'text' => 'New content',
);
$xlsx->addCell($content, 'B5', array(), array('useCellStyles' => true));

$xlsx->saveXlsx('example_addCell_7');