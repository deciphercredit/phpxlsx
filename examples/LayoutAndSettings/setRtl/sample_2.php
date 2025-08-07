<?php
// add a new sheet enabling rtl to a single sheet in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();
$xlsx->addSheet(array('rtl' => true));
$xlsx->addSheet();

$xlsx->saveXlsx('example_setRtl_2');