<?php
// add a remote image in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$options = array(
    'colSize' => 16,
    'rowSize' => 11,
);
$xlsx->addImage('http://www.2mdc.com/PHPDOCX/logo_badge.png', 'D6', $options);

$xlsx->saveXlsx('example_addImage_2');