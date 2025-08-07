<?php
// replace text variables (placeholders) with new text contents using ${} to wrap placeholders

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates_symbols.xlsx');
$xlsx->setTemplateSymbol('${', '}');

// replace variables in the sheets target
$xlsx->replaceVariableText(array('VAR_1' => 'New value 1', 'VAR_2' => 'New value 2', 'VAR_SHEET_2' => 'internal descr', 'VAR_SHEET_2_B' => 'info content', 'IMAGE_TITLE' => 'Title image'), array('target' => 'sheets'));

// replace variables in the footers and headers targets
$xlsx->replaceVariableText(array('VAR_HEADER_2' => 'Header content', 'HEADERS_1_1' => 'Header A', 'HEADERS_1_2' => 'Header B'), array('target' => 'headers'));
$xlsx->replaceVariableText(array('VAR_FOOTER_2' => 'Footer content', 'FOOTER_1_3' => 'Footer A', 'FOOTER_3_3' => 'Footer C'), array('target' => 'footers'));

// replace variables in the comments target
$xlsx->replaceVariableText(array('VAR_COMMENT' => 'More content in the comment'), array('target' => 'comments'));

// replace placeholder wrapped using #
$xlsx->setTemplateSymbol('#');
$xlsx->replaceVariableText(array('VAR_1' => 'New value #'));

$xlsx->saveXlsx('example_replaceVariableText_2');