<?php
// remove variables (placeholders)

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates.xlsx');

// remove variables in the sheets target
$xlsx->removeVariableText(array('VAR_1', 'VAR_2'), array('target' => 'sheets'));

// remove variables in the footers and headers targets
$xlsx->removeVariableText(array('VAR_HEADER_2', 'HEADERS_1_1', 'HEADERS_1_2'), array('target' => 'headers'));
$xlsx->removeVariableText(array('VAR_FOOTER_2', 'FOOTER_1_3', 'FOOTER_3_3'), array('target' => 'footers'));

// remove variables in the comments target
$xlsx->removeVariableText(array('VAR_COMMENT'), array('target' => 'comments'));

$xlsx->saveXlsx('example_removeVariableText_1');