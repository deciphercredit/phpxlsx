<?php
// get existing cell positions

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates.xlsx');

$cellPositions = $xlsx->getCellPositions();
print_r($cellPositions);