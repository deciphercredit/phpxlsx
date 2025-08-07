<?php
// modify the page layout of the active sheet setting a preset paper type and custom margins in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();
$xlsx->setSheetSettings(array('paperType' => 'A4', 'marginTop' => 0.1, 'marginBottom' => 0.1));

$xlsx->saveXlsx('example_setSheetSettings_1');