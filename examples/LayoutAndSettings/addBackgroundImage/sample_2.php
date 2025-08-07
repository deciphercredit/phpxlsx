<?php
// add a background image to second (active sheet in data_excel.xlsx) and third sheets in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');
$xlsx->addBackgroundImage('../../files/image.png');

$xlsx->setActiveSheet(array('position' => 2));

$xlsx->addBackgroundImage('../../files/image.png');

$xlsx->saveXlsx('example_addBackgroundImage_2');