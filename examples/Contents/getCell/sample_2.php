<?php
// get cells information from an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates.xlsx');

// get cell information
print_r($xlsx->getCell('A6'));

// get cell information
print_r($xlsx->getCell('C3'));