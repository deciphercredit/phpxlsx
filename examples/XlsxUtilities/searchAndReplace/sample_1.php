<?php
// replace string values

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Utilities\XlsxUtilities();

$data = array('B3' => 'other cell', 'merged Sheet 2' => 'merged cell');

$xlsx->searchAndReplace('../../files/data_excel.xlsx', 'example_searchAndReplace_1.xlsx', $data, 'sharedStrings');