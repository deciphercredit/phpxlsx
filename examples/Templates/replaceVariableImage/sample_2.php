<?php
// replace image variables (placeholders) in headers using a stream source. The placeholder has been added to the alt text content

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/templates.xlsx');

$xlsx->replaceVariableImage(array('IMAGE_1' => 'http://www.2mdc.com/PHPDOCX/logo_badge.png'));

$xlsx->saveXlsx('example_replaceVariableImage_2');