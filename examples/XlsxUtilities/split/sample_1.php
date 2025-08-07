<?php
// split an XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Utilities\XlsxUtilities();
$xlsx->split('../../files/data_excel.xlsx', 'splitXlsx_.xlsx');