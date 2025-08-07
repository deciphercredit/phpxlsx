<?php
// add HTML contents in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$html = '<p>Lorem ipsum</p>';
$xlsx->addHtml($html, 'A1');

$html = '<style>
    p {
        font-family: Arial;
        font-style: italic;
        text-decoration: underline;
        color: #2CA87A;
        vertical-align: bottom;
    }
</style>
<p>Lorem ipsum with Arial font</p>';
$xlsx->addHtml($html, 'A3');

$html = '<a href="https://www.phpxlsx.com">Link to phpxlsx</a>';
$xlsx->addHtml($html, 'B5');

$html = '<style>
    p {
        background-color: #AEBF7A;
    }
    .c1 {
        text-decoration: underline;
        font-size: 16pt;
    }
</style>
<p>
    <b>phpxlsx</b> can transform <span class="c1">HTML to XLSX</span>.
</p>
';
$xlsx->addHtml($html, 'A10');

$xlsx->saveXlsx('example_addHtml_1');