<?php
namespace Phpxlsx\Elements;
/**
 * Create html
 *
 * @category   Phpxlsx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class CreateHtml extends CreateElement
{
    /**
     *
     * @access protected
     * @var bool
     */
    protected $blockOpen;

    /**
     *
     * @access protected
     * @var array
     */
    protected $cellContents;

    /**
     *
     * @access protected
     * @var array
     */
    protected $cellStyles;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $hyperlinkOpen;

    /**
     *
     * @access protected
     * @var array
     */
    protected $position;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $tableOpen;

    /**
     *
     * @access protected
     * @var int
     */
    protected $tableOpenDepth;

    /**
     *
     * @access protected
     * @var array
     */
    protected $tableStyles;

    /**
     *
     * @access protected
     * @var array
     */
    protected $tableTr;

    /**
     *
     * @access protected
     * @var array
     */
    protected $tableTh;

    /**
     *
     * @access protected
     * @var array
     */
    protected $tableTd;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $thOpen;

    /**
     *
     * @access protected
     * @var CreateXlsx
     */
    protected $xlsx;

    /**
     * Constructor
     *
     * @param CreateXlsx $xlsx
     */
    public function __construct($xlsx)
    {
        $this->blockOpen = false;
        $this->cellContents = array();
        $this->cellStyles = array();
        $this->hyperlinkOpen  = false;
        $this->tableOpen  = false;
        $this->tableOpenDepth  = null;
        $this->tableStyles = array();
        $this->tableTr = array();
        $this->tableTh = array();
        $this->thOpen  = false;
        $this->tableTd = array();
        $this->xlsx = $xlsx;
    }

    /**
     * Create and add HTML
     *
     * @access public
     * @param string $html HTML to add
     *      <a>
     *      <p>, <h1>, <h2>, <h3>, <h4>, <h5>, <h6> : background-color, border (color, width), dir (ltr, rtl), text-align (left, center, right) vertical-align (top, center, bottom)
     *      <table>, <tr>, <th>, <td>               : background-color, dir (ltr, rtl), text-align (left, center, right) vertical-align (top, center, bottom)
     *      <span>, #text                           : color, font-family, font-size, font-style (italic, oblique), font-weight (bold, bolder, 700, 800, 900), text-decoration (line-through, underline), vertical-align (sub, super)
     *      <b>, <cite>, <em>, <i>, <s>, <strong>, <sub>, <sup>, <var>
     * @param array $position
     * @param array $options
     *      'baseURL' (string) Default as empty
     *      'disableWrapValue' (bool) if true disable using a wrap value with Tidy. Default as false
     *      'forceNotTidy' (bool) if true, avoid using Tidy. Only recommended if Tidy can't be installed. Default as false
     *      'insertMode' (string) replace, ignore. Default as replace
     *      'useHTMLExtended' (bool)  if true uses HTML extended tags and CSS extended styles. Default as false
     */
    public function createElementHtml($html, $position, $options)
    {
        $this->position = $position;

        require_once dirname(__FILE__) . '/../Libs/DOMPDF_lib.php';

        $dompdf = new \Phpxlsx\Libs\PARSERHTML();
        $dompdfTree = $dompdf->getDompdfTree($html, $options);

        $this->renderHtml($dompdfTree);

        // add remaining content if some exists
        $this->addCellContent();
    }

    /**
     * Render HTML
     *
     * @param array $node
     * @param int $depth
     */
    protected function renderHtml($node, $depth = 0)
    {
        $nodeAttributes = isset($node['attributes']) ? $node['attributes'] : array();
        $nodeProperties = isset($node['properties']) ? $node['properties'] : array();

        switch ($node['nodeName']) {
            case 'a':
            case 'p':
            case 'h1':
            case 'h2':
            case 'h3':
            case 'h4':
            case 'h5':
            case 'h6':
                // block content

                // add previous block contents to the XLSX
                $this->addCellContent($depth);

                // add the styles
                $this->cellStyles = array();
                if (isset($nodeProperties['background_color']) && is_array($nodeProperties['background_color'])) {
                    // backgroundColor
                    $color = $this->normalizeColor($nodeProperties['background_color']);

                    $this->cellStyles['backgroundColor'] = $color;
                }

                foreach (array('top', 'left', 'bottom', 'right') as $borderValue) {
                    // border
                    if (isset($nodeProperties['border_' . $borderValue . '_style']) && $nodeProperties['border_' . $borderValue . '_style'] != 'none') {
                        // border style
                        $borderStyle = $this->normalizeBorderStyle($nodeProperties['border_' . $borderValue . '_style']);
                        $borderScope = 'border' . ucwords($borderValue);

                        $this->cellStyles[$borderScope] = $borderStyle;

                        if (isset($nodeProperties['border_' . $borderValue . '_color'])) {
                            // border color
                            $borderColor = $this->normalizeColor($nodeProperties['border_' . $borderValue . '_color']);
                            $borderColorScope = 'borderColor' . ucwords($borderValue);

                            $this->cellStyles[$borderColorScope] = $borderColor;
                        }
                    }
                }
                if (isset($nodeAttributes['dir']) && !empty($nodeAttributes['dir'])) {
                    // direction
                    $this->cellStyles['textDirection'] = $nodeAttributes['dir'];
                }
                if (isset($nodeProperties['text_align'])) {
                    // textAlign
                    $horizontalAlign = $this->normalizeHorizontalAlign($nodeProperties['text_align']);

                    $this->cellStyles['horizontalAlign'] = $horizontalAlign;
                }
                if (isset($nodeProperties['vertical_align'])  && $nodeProperties['vertical_align'] != 'baseline') {
                    // verticalAlign
                    $verticalAlign = $this->normalizeVerticalAlign($nodeProperties['vertical_align']);

                    $this->cellStyles['verticalAlign'] = $verticalAlign;
                }

                // handle hyperlink content
                if ($node['nodeName'] == 'a') {
                    $this->hyperlinkOpen = true;
                    if (isset($nodeAttributes['href']) && !empty($nodeAttributes['href'])) {
                        $this->cellStyles['hyperlinkHref'] = $nodeAttributes['href'];
                    }
                }
                break;
            case 'img':
                break;
            case 'table':
                // add previous block contents to the XLSX
                $this->addCellContent($depth);

                $this->tableOpen = true;
                $this->tableOpenDepth = $depth;
                $this->tableStyles = array();

                if (isset($nodeProperties['data_style'])) {
                    $this->tableStyles['tableStyle'] = $nodeProperties['data_style'];
                }

                // new contents will be added
                $this->tableTr = array();
                $this->tableTd = array();
                $this->tableTh = array();
                break;
            case 'tr':
                // keep contents per row
                if (count($this->tableTd) > 0) {
                    $this->tableTr[count($this->tableTr) - 1] = $this->tableTd;
                }

                $this->tableTr[] = array();

                // new columns will be added
                $this->tableTd = array();
                break;
            case 'th':
            case 'td':
                // keep contents per column
                if ($node['nodeName'] == 'th') {
                    // th tags are added as table headers
                    $this->tableTh[] = array();
                    $this->thOpen  = true;
                } else {
                    // td tags are table as table contents
                    $this->tableTd[] = array();
                    $this->thOpen  = false;
                }

                // keep styles to be applied to the cell
                $this->cellStyles = array();
                if (isset($nodeProperties['background_color']) && is_array($nodeProperties['background_color'])) {
                    // backgroundColor
                    $color = $this->normalizeColor($nodeProperties['background_color']);

                    $this->cellStyles['backgroundColor'] = $color;
                }
                if (isset($nodeAttributes['dir']) && !empty($nodeAttributes['dir'])) {
                    // direction
                    $this->cellStyles['textDirection'] = $nodeAttributes['dir'];
                }
                if (isset($nodeProperties['text_align'])) {
                    // textAlign
                    $horizontalAlign = $this->normalizeHorizontalAlign($nodeProperties['text_align']);

                    $this->cellStyles['horizontalAlign'] = $horizontalAlign;
                }
                if (isset($nodeProperties['vertical_align'])  && $nodeProperties['vertical_align'] != 'baseline') {
                    // verticalAlign
                    $verticalAlign = $this->normalizeVerticalAlign($nodeProperties['vertical_align']);

                    $this->cellStyles['verticalAlign'] = $verticalAlign;
                }
                break;
            case '#text':
                // add the text contents
                $cellContents = array();
                if (isset($nodeProperties['color']) && !empty($nodeProperties['color']) && is_array($nodeProperties['color'])) {
                    // color
                    $color = $this->normalizeColor($nodeProperties['color']);

                    $cellContents['color'] = $color;
                }
                if (isset($nodeProperties['font_family']) && $nodeProperties['font_family'] != 'serif') {
                    // font
                    $arrayFonts = explode(',', $nodeProperties['font_family']);
                    $font = trim($arrayFonts[0]);
                    $font = str_replace(array('"', "'"), '', $font);

                    $cellContents['font'] = $font;
                }
                if (isset($nodeProperties['font_size']) && !empty($nodeProperties['font_size'])) {
                    // fontSize
                    $cellContents['fontSize'] = round($nodeProperties['font_size']);
                }
                if (isset($nodeProperties['font_style']) && ($nodeProperties['font_style'] == 'italic' || $nodeProperties['font_style'] == 'oblique')) {
                    // italic
                    $cellContents['italic'] = true;
                }
                if (isset($nodeProperties['font_weight']) && ($nodeProperties['font_weight'] == 'bold' || $nodeProperties['font_weight'] == 'bolder' || $nodeProperties['font_weight'] == '700' || $nodeProperties['font_weight'] == '800' || $nodeProperties['font_weight'] == '900')) {
                    // bold
                    $cellContents['bold'] = true;
                }
                if (isset($nodeProperties['text_decoration']) && $nodeProperties['text_decoration'] == 'line-through') {
                    // strikethrough
                    $cellContents['strikethrough'] = 'single';
                }
                if (isset($nodeProperties['text_decoration']) && $nodeProperties['text_decoration'] == 'underline') {
                    // underline
                    $cellContents['underline'] = 'single';
                }
                if (isset($nodeProperties['vertical_align']) && $nodeProperties['vertical_align'] == 'sub') {
                    // sub
                    $cellContents['subscript'] = true;
                }
                if (isset($nodeProperties['vertical_align']) && $nodeProperties['vertical_align'] == 'super') {
                    // super
                    $cellContents['superscript'] = true;
                }

                if (isset($node['nodeValue'])) {
                    // text content
                    $cellContents['text'] = $node['nodeValue'];
                }

                if ($this->tableOpen) {
                    if (!isset($cellContents['text'])) {
                        // avoid empty cells
                        $cellContents['text'] = '';
                    }

                    $cellContents['cellStyles'] = $this->cellStyles;
                    if ($this->thOpen) {
                        // add the cell contents into the last position of the th
                        $this->tableTh[count($this->tableTh) - 1][] = $cellContents;
                    } else {
                        // add the cell contents into the last position of the td
                        $this->tableTd[count($this->tableTd) - 1][] = $cellContents;
                    }

                    $this->cellStyles = array();
                } else {
                    $this->cellContents[] = $cellContents;
                }

                break;
            case 'close':
                // add previous block contents to the XLSX
                $this->addCellContent($depth);
            default:
                break;
        }

        $depth++;

        if (isset($nodeProperties['display']) && $nodeProperties['display'] == 'none') {
            //do not render the subtree
        } else {
            if (!empty($node['children'])) {
                foreach ($node['children'] as $child) {
                    $this->renderHtml($child, $depth);
                }
            }
        }
    }

    /**
     * Adds cell content
     *
     * @access private
     * @param int $depth
     */
    private function addCellContent($depth = 0)
    {
        if ($this->tableOpen && !is_null($this->tableOpenDepth) && $this->tableOpenDepth == $depth) {
            $optionsTable = array();

            if (count($this->tableTd) > 0) {
                // get remaining rows
                $this->tableTr[count($this->tableTr) - 1] = $this->tableTd;
            }

            // handle tables
            $tableContents = $this->tableTr;
            foreach ($tableContents as $tableKey => $tableContent) {
                // remove empty rows to not add them as table contents
                if (count($tableContent) == 0) {
                    unset($tableContents[$tableKey]);
                }
            }
            // reorder array, needed by addTable to start from 0
            $tableContents = array_values($tableContents);

            // get column names
            if (count($this->tableTh) > 0) {
                $columnNames = array();
                foreach ($this->tableTh as $tableThContent) {
                    $columnNames[] = $tableThContent;
                }
                $optionsTable['columnNames'] = $columnNames;
            }

            $positionTable = $this->xlsx->addTable($tableContents, $this->position['text'] . $this->position['number'], $this->tableStyles, $optionsTable);

            // set the new position after adding the table
            $numberPosition = preg_replace('/[^0-9]/', '', $positionTable['to']);
            $numberPosition++;
            $this->position['number'] = $numberPosition;

            // empty cell contents and styles
            $this->cellStyles = array();
            $this->cellContents = array();

            $this->tableOpen = false;
            $this->tableOpenDepth = null;
            $this->tableStyles = array();
            $this->tableTr = array();
            $this->tableTh = array();
            $this->tableTd = array();
        } else {
            // handle cell contents
            if (count($this->cellStyles) > 0 || count($this->cellContents) > 0 || $this->hyperlinkOpen) {
                if ($this->hyperlinkOpen) {
                    // handle link content
                    if (isset($this->cellStyles['hyperlinkHref'])) {
                        $this->xlsx->addLink($this->cellStyles['hyperlinkHref'], $this->position['text'] . $this->position['number'], $this->cellContents, $this->cellStyles);
                    }
                    $this->hyperlinkOpen = false;
                } else {
                    // handle cell content
                    $this->xlsx->addCell($this->cellContents, $this->position['text'] . $this->position['number'], $this->cellStyles);
                }

                (int)$this->position['number']++;
                $this->cellStyles = array();
                $this->cellContents = array();
            }
        }
    }

    /**
     * Normalizes a border style
     *
     * @access private
     * @param string $borderStyle Border style
     * @return string
     */
    private function normalizeBorderStyle($borderStyle)
    {
        $borderStylesHTML = array(
            'none' => 'none',
            'dotted' => 'dotted',
            'dashed' => 'dashed',
            'solid' => 'thin',
            'double' => 'double',
            'groove' => 'threeDEngrave',
            'ridge' => 'thin',
        );

        if (!empty($borderStyle)) {
            if (array_key_exists($borderStyle, $borderStylesHTML)) {
                return $borderStylesHTML[$borderStyle];
            } else {
                return 'thin';
            }
        } else {
            return '';
        }
    }

    /**
     * Normalizes a color
     *
     * @access private
     * @param array $color
     * @return string
     */
    private function normalizeColor($color)
    {
        if (strtolower($color['hex']) == 'transparent') {
            return '';
        } else {
            return strtoupper(str_replace('#', '', $color['hex']));
        }
    }

    /**
     * Normalizes a horizontal align value
     *
     * @access private
     * @param string $horizontalAlign Horizontal align
     * @return string
     */
    private function normalizeHorizontalAlign($horizontalAlign = 'left')
    {
        $normalizedHorizontalAlign = $horizontalAlign;

        switch ($horizontalAlign) {
            case 'left':
            case 'justify':
                $normalizedHorizontalAlign = 'left';
                break;
            case 'center':
                $normalizedHorizontalAlign = 'center';
                break;
            case 'right':
                $normalizedHorizontalAlign = 'right';
                break;
            default:
                $normalizedHorizontalAlign = 'left';
        }

        return $normalizedHorizontalAlign;
    }

    /**
     * Normalizes a vertical align value
     *
     * @access private
     * @param string $verticalAlign Vertical align
     * @return string
     */
    private function normalizeVerticalAlign($verticalAlign = 'baseline')
    {
        $normalizedVerticalAlign = $verticalAlign;

        switch ($verticalAlign) {
            case 'super':
            case 'top':
            case 'text-top':
                $normalizedVerticalAlign = 'top';
                break;
            case 'center':
            case 'middle':
                $normalizedVerticalAlign = 'center';
                break;
            case 'sub':
            case 'baseline':
            case 'bottom':
            case 'text-bottom':
            default:
                $normalizedVerticalAlign = 'bottom';
        }

        return $normalizedVerticalAlign;
    }
}