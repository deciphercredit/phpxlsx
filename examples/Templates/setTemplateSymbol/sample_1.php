<?php
// change the symbol used to wrap variables (placehoders) and replace text variables (placeholder)

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates_symbols.xlsx');

// replace variables with #
$xlsx->setTemplateSymbol('#');
$xlsx->replaceVariableText(array('VAR_1' => 'New value 1'), array('target' => 'sheets'));

// replace variables with ${}
$xlsx->setTemplateSymbol('${', '}');
$xlsx->replaceVariableText(array('VAR_1' => 'New value 2'), array('target' => 'sheets'));

$xlsx->saveXlsx('example_setTemplateSymbol_1');