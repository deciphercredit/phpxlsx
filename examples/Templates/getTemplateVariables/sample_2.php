<?php
// return the variables (placeholders) using ${} to wrap placeholders

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates_symbols.xlsx');
$xlsx->setTemplateSymbol('${', '}');

print_r($xlsx->getTemplateVariables());

$xlsx->setTemplateSymbol('#');

print_r($xlsx->getTemplateVariables());