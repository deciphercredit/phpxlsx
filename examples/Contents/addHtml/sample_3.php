<?php
// add HTML tables in an XLSX created from scratch

require_once dirname( __FILE__ ) . '/../../../Classes/Phpxlsx/Create/CreateXlsx.php';

$xlsx = new Phpxlsx\Create\CreateXlsx();

$html = '
<style>
    .cstable2 {
        data-style: TableStyleMedium2;
    }
    tr.rowstyles {
        font-family: Arial;
        font-weight: bold;
        text-decoration: underline;
        color: #2CA87A;
    }
</style>
<table class="cstable2">
    <tbody>
        <tr>
            <td style="background-color: yellow;">Cell 1 1</td>
            <td>Cell 1 2</td>
        </tr>
        <tr>
            <td>Cell 2 1</td>
            <td></td>
        </tr>
        <tr class="rowstyles">
            <td><em>Cell</em> 3 1</td>
            <td><em>Cell</em> 3 2</td>
        </tr>
    </tbody>
</table>
<p>Paragraph after table.</p>
';
$xlsx->addHtml($html, 'B3');

// table with th tags, that are added as table headers
$html = '
<style>
    tr.rowstyles {
        font-family: Arial;
        font-weight: bold;
        text-decoration: underline;
        color: #2CA87A;
    }
</style>
<table>
    <tbody>
        <tr>
            <th>Col A</th>
            <th>Col B</th>
        </tr>
        <tr>
            <td>Cell 2 1</td>
            <td>Cell 2 2</td>
        </tr>
        <tr class="rowstyles">
            <td><em>Cell</em> 3 1</td>
            <td><em>Cell</em> 3 2</td>
        </tr>
    </tbody>
</table>
';
$xlsx->addHtml($html, 'G3');

$xlsx->saveXlsx('example_addHtml_3');