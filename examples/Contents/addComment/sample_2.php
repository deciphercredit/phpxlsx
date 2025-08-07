<?php
// add comments in an existing XLSX

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsxFromTemplate('../../files/data_excel.xlsx');

// add a comment with an author
$xlsx->addComment('My comment', 'B5', array('author' => 'phpxlsx'));

// add a comment with styles
$comment = array(
    'text' => 'Comment with styles',
    'bold' => true,
    'italic' => true,
    'underline' => 'single',
);
$xlsx->addComment($comment, 'E5', array('author' => 'phpxlsx'));

$xlsx->saveXlsx('example_addComment_2');