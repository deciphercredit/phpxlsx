<?php
// set custom row height values in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
);
$xlsx->addCell($content, 'A2');

// set row settings. The position is set using row numbers
$xlsx->setRowSettings(4, array('height' => 90.25));
$xlsx->setRowSettings(2, array('height' => 60));
$xlsx->setRowSettings(3, array('height' => 50.25));

$xlsx->addSheet(array('name'=> 'Other', 'removeSelected' => true, 'selected' => true, 'active' => true));

$content = array(
    'text' => 'Sed ut perspiciatis unde omnis',
);
$xlsx->addCell($content, 'A1');

// set row settings. The position is set using cell positions
$xlsx->setRowSettings('A1', array('height' => 40));
$xlsx->setRowSettings('B2', array('height' => 50));
$xlsx->setRowSettings('B3', array('height' => 40));

$xlsx->saveXlsx('example_setRowSettings_1');