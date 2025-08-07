<?php
// get the active sheet in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates.xlsx');
$activeSheet = $xlsx->getActiveSheet();

var_dump($activeSheet);