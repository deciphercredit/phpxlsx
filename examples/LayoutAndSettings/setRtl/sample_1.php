<?php
// add a new sheet enabling rtl as global option in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();
$xlsx->setSheetSettings(array('rtl' => true));
$xlsx->setRtl();

$xlsx->addSheet();

$xlsx->saveXlsx('example_setRtl_1');