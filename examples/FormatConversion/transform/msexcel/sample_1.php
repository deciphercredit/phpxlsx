<?php
// transform an XLSX to PDF using the conversion plugin based on MS Excel

require_once dirname( __FILE__ ) . '/../../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

// global paths must be used
$xlsx->transform('E:\\VMs\\shared\\phpdocx\\repos\\phpxlsx\\examples\\files\\data_excel.xlsx', 'E:\\VMs\\shared\\phpdocx\\repos\\phpxlsx\\examples\\FormatConversion\\transform\\msexcel\\transform_msexcel_1.pdf', 'msexcel');