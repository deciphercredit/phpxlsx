<?php
// create and apply multiple cell styles in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

// create the new cell styles
$xlsx->createCellStyle('MyStyle1', array('bold' => true, 'fontSize' => 14, 'backgroundColor' => 'FFFF00'));
$xlsx->createCellStyle('My Style 2', array('italic' => true, 'fontSize' => 7, 'backgroundColor' => 'FF0000'));

$content = array(
    'text' => 'Style 1',
);
$xlsx->addCell($content, 'B12', array('cellStyleName' => 'MyStyle1'));
$content = array(
    'text' => 'Style 2',
);
$xlsx->addCell($content, 'B14', array('cellStyleName' => 'My Style 2'));
$content = array(
    'text' => 'Style 3',
);
$xlsx->addCell($content, 'B16', array('cellStyleName' => 'MyStyle1'));
$content = array(
    'text' => 'Good',
);
$xlsx->addCell($content, 'B18', array('cellStyleName' => 'Good'));

$xlsx->saveXlsx('example_createCellStyle_4');