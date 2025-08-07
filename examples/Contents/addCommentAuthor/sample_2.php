<?php
// add comment authors in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

$xlsx->addCommentAuthor('phpxlsxauthor');
$xlsx->addCommentAuthor('myauthor');

$xlsx->saveXlsx('example_addCommentAuthor_2');