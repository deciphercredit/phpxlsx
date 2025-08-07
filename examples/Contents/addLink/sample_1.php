<?php
// add URL link contents in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$content = array(
    'text' => 'Link to phpxlsx',
);
$xlsx->addLink('https://www.phpxlsx.com', 'B2', $content);

$content = array(
    'text' => 'Link to phpdocx',
    'bold' => true,
    'color' => '000000',
    'underline' => 'none',
);
$xlsx->addLink('https://www.phpdocx.com', 'D5', $content);

$xlsx->saveXlsx('example_addLink_1');