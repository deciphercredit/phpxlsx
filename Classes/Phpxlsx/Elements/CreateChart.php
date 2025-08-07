<?php
namespace Phpxlsx\Elements;
/**
 * Create chart
 *
 * @category   Phpxlsx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class CreateChart extends CreateElement
{
    /**
     * Create chart
     *
     * @access public
     * @param string $chart
     * @param array $position
     * @param array $chartOptions
     *      'axPos' (array) position of the axis (r, l, t, b), each value of the array for each position (if a value if null avoids adding it)
     *      'border' (int) border width in points
     *      'colOffset' (array) given in emus (1cm = 360000 emus). 0 as default
     *          'from' (int) from offset
     *          'to' (int) from offset
     *      'colSize' (int) number of cols used by the chart
     *      'color' (string) (1, 2, 3...) color scheme
     *      'comboChart' chart to add as combo chart. Use with the returnChart option. Global styles and properties are shared with the base chart. For bar, col, line, area, and radar charts
     *      'data' (array) values
     *      'font' (string) Arial, Times New Roman ...
     *      'formatCode' (string) number format
     *      'formatDataLabels' (array)
     *          'rotation' => (int)
     *          'position' => (string) center, insideEnd, insideBase, outsideEnd
     *      'haxLabel' (bool) horizontal axis label
     *      'haxLabelDisplay' (string) rotated, vertical, horizontal
     *      'hgrid' (int) 0 (no grid) 1 (only major grid lines - default) 2 (only minor grid lines) 3 (both major and minor grid lines)
     *      'legendOverlay' (bool) if true the legend may overlay the chart
     *      'legendPos' (String) r, l, t, b, none
     *      'majorUnit' (float) bar, col, line charts
     *      'minorUnit' (float) bar, col, line charts
     *      'orientation' (array) orientation of the axis, from min to max (minMax) or max to min (maxMin), each value of the array for each axis (if a value if null avoids adding it)
     *      'rowOffset' (array) given in emus (1cm = 360000 emus). 0 as default
     *          'from' (int) from offset
     *          'to' (int) from offset
     *      'rowSize' (int) number of rows used by the chart
     *      'scalingMax' (float) scaling max value bar, col, line charts
     *      'scalingMin' (float) scaling min value bar, col, line charts
     *      'showCategory' (bool) shows the categories inside the chart
     *      'showLegendKey' (bool) if true shows the legend values
     *      'showPercent' (bool) if true shows the percent values
     *      'showSeries' (bool) if true shows the series values
     *      'showTable' (bool) if true shows the table of values
     *      'showValue' (bool) if true shows the values inside the chart
     *      'stylesTitle' (array)
     *          'bold' (bool)
     *          'color' (string) FFFFFF, FF0000
     *          'font' (string)  Arial, Times New Roman ...
     *          'fontSize' (int) 8, 9, 10, ... size as drawing content (10 to 400000). 1420 as default
     *          'italic' (bool)
     *      'tickLblPos' (mixed) tick label position (nextTo, high, low, none). If string, uses default values. If array, sets a value for each position
     *      'title' (string)
     *      'trendline' (array of trendlines). Compatible with line, bar and col 2D charts
     *          'color' (string) 0000FF
     *          'displayEquation' (bool) display equation on chart
     *          'displayRSquared' (bool) display R-squared value on chart
     *          'intercept' (float) set intercept
     *          'lineStyle' (string) solid, dot, dash, lgDash, dashDot, lgDashDot, lgDashDotDot, sysDash, sysDot, sysDashDot, sysDashDotDot
     *          'type' (string) 'exp', 'linear', 'log', 'poly', 'power', 'movingAvg'
     *          'typeOrder' (int) for poly and movingAvg types
     *      'vaxLabel' (bool) vertical axis label
     *      'vaxLabelDisplay' (string) rotated, vertical, horizontal
     *      'vgrid'  (int) 0 (no grid) 1 (only major grid lines - default) 2 (only minor grid lines) 3 (both major and minor grid lines)
     *
     *  3D charts:
     *      'perspective' (int) 20, 30...
     *      'rotX' (int) 20, 30...
     *      'rotY' (int) 20, 30...
     *
     *  Bar and column charts:
     *      'gapWidth' (int) gap width
     *      'groupBar' (string) clustered, stacked, percentStacked
     *      'overlap' (int) overlap value
     *
     *  Line charts:
     *      'smooth' (mixed) enable smooth lines, line charts. '0' forces disabling it
     *      'symbol' (string) Line charts: none, dot, plus, square, star, triangle, x, diamond, circle and dash
     *      'symbolSize' (int) the size of the symbols (values 1 to 73)
     *
     *  Pie and doughnut charts:
     *      'explosion' (int) distance between the diferents values
     *      'holeSize' (int) size of the hole in doughnut type
     *
     *  Theme:
     *  'theme' (array):
     *      'chartArea' (array):
     *          'backgroundColor' (string)
     *      'gridLines' (array):
     *          'capType' (string)
     *          'color' (string): RGB
     *          'dashType' (string)
     *          'width' (int)
     *      'horizontalAxis' (array):
     *          'textBold' (bool)
     *          'textDirection' (string): 'horizontal', 'rotate90', 'rotate270'
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     *      'legendArea' (array):
     *          'backgroundColor' (string)
     *          'textBold' (bool)
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     *      'plotArea' (array):
     *          'backgroundColor' (string)
     *      'serDataLabels' (array): data labels options (bar, bubble, column, line ofPie, pie and scatter charts)
     *          'formatCode' (array)
     *          'position (array): center, insideEnd, insideBase, outsideEnd
     *          'showCategory' (array): 0, 1
     *          'showLegendKey' (array): 0, 1
     *          'showPercent' (array): 0, 1
     *          'showSeries' (array): 0, 1
     *          'showValue' (array): 0, 1
     *      'serRgbColors' (array): series colors
     *      'valueRgbColors' (array): values colors
     *      'verticalAxis' (array):
     *          'textBold' (bool)
     *          'textDirection' (string): 'horizontal', 'rotate90', 'rotate270'
     *          'textItalic' (bool)
     *          'textSize' (int): points
     *          'textUnderline' (string): DrawingML values such as 'none', 'sng', 'dash'
     * @param array $options
     *      'rId' (string) chart ID
     * @return array
     */
    public function createElementChart($chart, $position, $chartOptions = array(), $options = array())
    {
        // from and to column values
        $colFromValue = $this->wordToInt($position['text']);
        $colToValue = $colFromValue + 1;
        if (isset($chartOptions['colSize'])) {
            $colToValue = $colFromValue + $chartOptions['colSize'];
        } else {
            // set enough size to display the chart
            $colToValue += 6;
        }
        $colFromOffsetValue = 0;
        if (isset($chartOptions['colOffset']) && isset($chartOptions['colOffset']['from'])) {
            $colFromOffsetValue = $options['colOffset']['from'];
        }
        $colToOffsetValue = 0;
        if (isset($chartOptions['colOffset']) && isset($chartOptions['colOffset']['to'])) {
            $colToOffsetValue = $chartOptions['colOffset']['to'];
        }
        // used to get the width size. Not needed by the chart. 0 value is used
        //$width = (int)(($colToValue - $colFromValue) * ((14.4/20) * 1530350));

        // from and to row values
        $rowFromValue = $position['number'];
        // start from 0
        $rowFromValue--;
        $rowToValue = $rowFromValue + 1;
        if (isset($chartOptions['rowSize'])) {
            $rowToValue = $rowFromValue + $chartOptions['rowSize'];
        } else {
            // set enough size to display the chart
            $rowToValue += 10;
        }
        $rowFromOffsetValue = 0;
        if (isset($chartOptions['rowOffset']) && isset($chartOptions['rowOffset']['from'])) {
            $rowFromOffsetValue = $chartOptions['rowOffset']['from'];
        }
        $rowToOffsetValue = 0;
        if (isset($chartOptions['rowOffset']) && isset($chartOptions['rowOffset']['to'])) {
            $rowToOffsetValue = $chartOptions['rowOffset']['to'];
        }
        // used to get the height size. Not needed by the chart. 0 value is used
        //$height = (int)(($rowToValue - $rowFromValue) * ((8.11/20) * 1530350));

        $newContents = array(
            'chartXml' => '',
            'drawingXml' => '',
        );

        $name = 'Chart ' . $options['rId'];
        if (isset($chartOptions['name'])) {
            $name = $this->parseAndCleanTextString($chartOptions['name']);
        }

        $cNvPrId = rand(999999, 999999999);

        // drawing content
        $newContents['drawingXml'] = '<xdr:twoCellAnchor editAs="oneCell" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:from><xdr:col>'.$colFromValue.'</xdr:col><xdr:colOff>'.$colFromOffsetValue.'</xdr:colOff><xdr:row>'.$rowFromValue.'</xdr:row><xdr:rowOff>'.$rowFromOffsetValue.'</xdr:rowOff></xdr:from><xdr:to><xdr:col>'.$colToValue.'</xdr:col><xdr:colOff>'.$colToOffsetValue.'</xdr:colOff><xdr:row>'.$rowToValue.'</xdr:row><xdr:rowOff>'.$rowToOffsetValue.'</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="'.$cNvPrId.'" name="'.$name.'"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId'.$options['rId'].'" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';

        // generate the chart class to be used from the chart type
        $classType = ucwords(str_replace(array('3D', 'Col'), array('', 'Bar'), ucwords($chart)));
        // remove subtype strings
        $classType = str_replace(array('Cylinder', 'Cone', 'Pyramid'), '', $classType);
        $options['type'] = $chart;
        $chartClass = 'Phpxlsx\Charts\CreateChart' . $classType;
        $chartType = new $chartClass();
        // chart content
        $chartXml = $chartType->createChart($chartOptions, $options);

        // theme chart
        if (isset($options['theme']) && is_array($options['theme']) && count($options['theme']) > 0) {
            $themeChart = new \Phpxlsx\Theme\ThemeCharts();
            $chartXml = $themeChart->theme($chartXml, $options['theme']);
        }

        $newContents['chartXml'] = $chartXml;

        return $newContents;
    }
}