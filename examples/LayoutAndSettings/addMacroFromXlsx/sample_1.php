<?php
// add a macro in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();
$xlsx->addMacroFromXlsx('../../files/sample_macro.xlsm');

$xlsx->saveXlsx('example_addMacroFromXlsx_1.xlsm');