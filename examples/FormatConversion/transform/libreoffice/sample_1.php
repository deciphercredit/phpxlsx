<?php
// transform an XLSX to PDF using the conversion plugin based on LibreOffice

require_once dirname( __FILE__ ) . '/../../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();
$xlsx->transform('../../../files/data_excel.xlsx', 'transform_libreoffice_1.pdf', 'libreoffice');