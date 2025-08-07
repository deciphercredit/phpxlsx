<?php
// add a CSV file

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$xlsx->addCsv('../../files/sample.csv', 'A1');

$xlsx->saveXlsx('example_addCsv_1');