<?php
// apply preset cell styles in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Good',
);
$xlsx->addCell($content, 'A1', array('cellStyleName' => 'Good'));

$content = array(
    'text' => 'Bad',
);
$xlsx->addCell($content, 'D3', array('cellStyleName' => 'Bad'));

$content = array(
    'text' => 'Error',
);
$xlsx->addCell($content, 'D6', array('cellStyleName' => 'Warning Text'));

$xlsx->saveXlsx('example_createCellStyle_2');