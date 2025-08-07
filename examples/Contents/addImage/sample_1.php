<?php
// add an image in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$options = array(
    'colSize' => 3,
    'rowSize' => 7,
);
$xlsx->addImage('../../files/image.png', 'A1', $options);

$options = array(
    'colSize' => 6,
    'colOffset' => array('from' => 129600, 'to' => 137160),
    'rowSize' => 11,
    'rowOffset' => array('from' => 152280, 'to' => 159840),
);
$xlsx->addImage('../../files/image.jpg', 'AB4', $options);

$xlsx->saveXlsx('example_addImage_1');