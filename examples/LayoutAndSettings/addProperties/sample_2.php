<?php
// add date properties in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$properties = array(
    'created' => '2016-11-21T09:00:00Z', // force a value
    'modified' => substr(date(DATE_W3C), 0, 19) . 'Z', // dynamic value
    'lastModifiedBy' => 'phpxlsxuser',
);
$xlsx->addProperties($properties);

$xlsx->saveXlsx('example_addProperties_2');