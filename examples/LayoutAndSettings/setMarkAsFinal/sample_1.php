<?php
// set mark as final to the document to prevent changing it in an XLSX created from scratch. Premium licenses include crypto and sign features to get better protection

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$xlsx->setMarkAsFinal();

$xlsx->saveXlsx('example_setMarkAsFinal_1');