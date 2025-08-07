<?php
// add a background image to all sheets in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();
$xlsx->addBackgroundImage('../../files/image.png');

$xlsx->saveXlsx('example_addBackgroundImage_1');