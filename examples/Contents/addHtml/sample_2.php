<?php
// add HTML contents in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$html = '
<style>
    #i1 {
        font-size: 18px;
    }
    .c1 {
        text-decoration: underline;
    }
    .c2 {
        color: #2CA87A;
    }
</style>
<p>
    <p>
        <b>phpxlsx</b> can transform <span class="c1">HTML to XLSX</span>.
    </p>
    <p id="i1" class="c2">
        <span class="c1">CSS styles</span> using ids and classes are supported.
    </p>
    <p class="c2">
        CSS styles classes are supported.
    </p>
</p>
<p>Links can be added:</p>
<a href="https://www.phpxlsx.com">Link</a>
<p>Heading tags are added as block elements:</p>
<h1>Heading 1</h1>
<h2>Heading 2</h2>
<h3>Heading 3</h3>
<h4>Heading 4</h4>
<h5>Heading 5</h5>
<h6>Heading 6</h6>
';
$xlsx->addHtml($html, 'A1');

$xlsx->saveXlsx('example_addHtml_2');