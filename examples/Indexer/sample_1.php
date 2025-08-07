<?php
// return information from an existing XLSX

require_once dirname( __FILE__ ) . '/../../Classes/Phpxlsx/Create/CreateXlsx.php';

$indexer = new Phpxlsx\Utilities\Indexer('../files/indexer.xlsx');
$output = $indexer->getOutput();

print_r('sheets: ');
print_r($output['sheets']);

print_r('workbooks: ');
print_r($output['workbooks']);

/*
print_r('comments: ');
print_r($output['comments']);
*/

/*
print_r('core properties: ');
print_r($output['properties']['core']);

print_r('custom properties: ');
print_r($output['properties']['custom']);
*/