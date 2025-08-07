<?php
// add a new sheet in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/template_sample.xlsx');

$xlsx->addSheet();

$xlsx->saveXlsx('example_addSheet_2');