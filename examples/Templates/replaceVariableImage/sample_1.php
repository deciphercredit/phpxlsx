<?php
// replace image variables (placeholders). The placeholders have been added to the alt text

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates.xlsx');

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

$xlsx->saveXlsx('example_replaceVariableImage_1');