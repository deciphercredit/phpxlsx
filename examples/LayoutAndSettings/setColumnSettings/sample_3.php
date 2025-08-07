<?php
// set hidden columns in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Lorem ipsum dolor sit amet',
);
$xlsx->addCell($content, 'A2');

// hide columns. The position is set using column letters
$xlsx->setColumnSettings('B', array('hidden' => true));
$xlsx->setColumnSettings('D', array('hidden' => true));
$xlsx->setColumnSettings('E', array('hidden' => true));

$content = array(
    'text' => 'Value',
);
$xlsx->addCell($content, 'E5');

// unhide a column
$xlsx->setColumnSettings('E', array('hidden' => false));
$xlsx->setColumnSettings('E', array('width' => 20));

$xlsx->saveXlsx('example_setColumnSettings_3');