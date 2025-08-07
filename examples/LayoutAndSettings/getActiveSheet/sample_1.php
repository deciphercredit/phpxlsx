<?php
// get the active sheet in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();
$activeSheet = $xlsx->getActiveSheet();

var_dump($activeSheet);

$xlsx->addSheet(array('name'=> 'Other', 'removeSelected' => true, 'selected' => true, 'active' => true));

$activeSheet = $xlsx->getActiveSheet();

var_dump($activeSheet);