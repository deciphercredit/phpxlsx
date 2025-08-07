<?php
// add an SVG image file

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$options = array(
    'colSize' => 3,
    'rowSize' => 7,
);
$xlsx->addSvg('../../files/image.svg', 'A1', $options);

$xlsx->saveXlsx('example_addSvg_1');