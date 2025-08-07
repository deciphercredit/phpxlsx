<?php
// add a defined name in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 100,
);
$xlsx->addCell($content, 'A2');
$content = array(
    'text' => 10.50,
);
$xlsx->addCell($content, 'A3');

$xlsx->addDefinedName('MY_SUM', 'Sheet1!$A$2:$A$3');

$contentStyles = array(
    'bold' => true,
    'font' => 'Arial',
);
$cellStyles = array(
    'backgroundColor' => 'FFFF00',
    'verticalAlign' => 'top',
);
$xlsx->addFunction('=SUM(MY_SUM)', 'D3', $contentStyles, $cellStyles);

$xlsx->saveXlsx('example_addDefinedName_1');