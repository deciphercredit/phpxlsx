<?php
// replace image variables (placeholders) using ${} to wrap placeholders. The placeholder has been added to the alt text content

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates_symbols.xlsx');
$xlsx->setTemplateSymbol('${', '}');

// replace image in sheets target
$xlsx->replaceVariableImage(array('IMAGE_1' => '../../files/image.png'));

// replace image in headers target
$options = array(
	'target' => 'headers',
);
$xlsx->replaceVariableImage(array('IMAGE_HEADER' => '../../files/image.jpg'), $options);

// replace image in footers target
$options = array(
	'target' => 'footers',
);
$xlsx->replaceVariableImage(array('IMAGE_FOOTER' => '../../files/image.jpg'), $options);

$xlsx->saveXlsx('example_replaceVariableImage_3');