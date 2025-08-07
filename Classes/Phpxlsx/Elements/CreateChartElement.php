<?php
namespace Phpxlsx\Elements;
/**
 * Create charts
 *
 * @category   Phpxlsx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class CreateChartElement extends CreateElement
{
    /**
     *
     * @access protected
     * @var int
     */
    protected $autoUpdate;

    /**
     *
     * @access protected
     * @var array
     */
    protected $axPos;

    /**
     *
     * @access protected
     * @var string
     */
    protected $border;

    /**
     *
     * @access protected
     * @var string
     */
    protected $color;

    /**
     *
     * @access protected
     * @var array
     */
    protected $custSplit;

    /**
     *
     * @access protected
     * @var array
     */
    protected $data;

    /**
     *
     * @access protected
     * @var int
     */
    protected $delete;

    /**
     *
     * @access protected
     * @var string
     */
    protected $explosion;

    /**
     *
     * @access protected
     * @var string
     */
    protected $font;

    /**
     *
     * @access protected
     * @var string
     */
    protected $formatCode;

    /**
     *
     * @access protected
     * @var array
     */
    protected $formatDataLabels;

    /**
     *
     * @access protected
     * @var int
     */
    protected $gapWidth;

    /**
     *
     * @access protected
     * @var string
     */
    protected $groupBar;

    /**
     *
     * @access protected
     * @var string
     */
    protected $haxLabel;

    /**
     *
     * @access protected
     * @var string
     */
    protected $haxLabelDisplay;

    /**
     *
     * @access protected
     * @var string
     */
    protected $hgrid;

    /**
     *
     * @access protected
     * @var string
     */
    protected $holeSize;

    /**
     *
     * @access protected
     * @var int
     */
    protected $horizontalOffset;

    /**
     *
     * @access protected
     * @var string
     */
    protected $legendOverlay;

    /**
     *
     * @access protected
     * @var string
     */
    protected $legendPos;

    /**
     * @access protected
     * @var float
     */
    protected $majorUnit;

    /**
     * @access protected
     * @var float
     */
    protected $minorUnit;

    /**
     *
     * @access protected
     * @var string
     */
    protected $name;

    /**
     *
     * @access protected
     * @var string
     */
    protected $options;

    /**
     *
     * @access protected
     * @var array
     */
    protected $orientation;

    /**
     *
     * @access protected
     * @var int
     */
    protected $overlap;

    /**
     *
     * @access protected
     * @var int
     */
    protected $perspective;

    /**
     *
     * @access protected
     * @var int
     */
    protected $rAngAx;

    /**
     *
     * @access protected
     * @var string
     */
    protected $rId;

    /**
     *
     * @access protected
     * @var int
     */
    protected $rotX;

    /**
     *
     * @access protected
     * @var int
     */
    protected $rotY;

    /**
     *
     * @access protected
     * @var int
     */
    protected $roundedCorners;

    /**
     * @access protected
     * @var float
     */
    protected $scalingMax;

    /**
     * @access protected
     * @var float
     */
    protected $scalingMin;

    /**
     * @access protected
     * @var float
     */
    protected $secondPieSize;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $showBubbleSize;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $showCategory;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $showLegendKey;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $showPercent;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $showSeries;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $showTable;

    /**
     *
     * @access protected
     * @var bool
     */
    protected $showValue;

    /**
     *
     * @access protected
     * @var boolean
     */
    protected $smooth;

    /**
     *
     * @access protected
     * @var int
     */
    protected $splitPos;

    /**
     *
     * @access protected
     * @var string
     */
    protected $splitType;

    /**
     *
     * @access protected
     * @var string
     */
    protected $style;

    /**
     *
     * @access protected
     * @var array
     */
    protected $stylesTitle;

    /**
     *
     * @access protected
     * @var string
     */
    protected $subtype;

    /**
     *
     * @access protected
     * @var string
     */
    protected $symbol;

    /**
     *
     * @access protected
     * @var string
     */
    protected $symbolSize;

    /**
     *
     * @access protected
     * @var string
     */
    protected $textalign;

    /**
     *
     * @access protected
     * @var array
     */
    protected $theme;

    /**
     * @access protected
     * @var string
     */
    protected $tickLblPos;

    /**
     *
     * @access protected
     * @var string
     */
    protected $title;

    /**
     *
     * @access protected
     * @var string
     */
    protected $type;

    /**
     *
     * @access protected
     * @var array
     */
    protected $values;

    /**
     *
     * @access protected
     * @var int
     */
    protected $varyColors;

    /**
     *
     * @access protected
     * @var string
     */
    protected $vaxLabel;

    /**
     *
     * @access protected
     * @var string
     */
    protected $vaxLabelDisplay;

    /**
     *
     * @access protected
     * @var int
     */
    protected $verticalOffset;

    /**
     *
     * @access protected
     * @var string
     */
    protected $vgrid;

    /**
     *
     * @access protected
     * @var boolean
     */
    protected $wireframe;

    /**
     *
     * @access protected
     * @var string
     */
    protected $xmlChart;

    /**
     * Construct
     *
     * @access public
     */
    public function __construct()
    {
        //set for 2010 compatibility
        $this->varyColors = 0;
        $this->autoUpdate = 0;
        $this->delete = 0; //removes the axis if set to 1
        $this->rAngAx = 0;
        $this->roundedCorners = 0;

        $this->rId = '';
        $this->type = '';
        $this->data = array();
        $this->rotX = '';
        $this->rotY = '';
        $this->perspective = '';
        $this->color = '';
        $this->groupBar = '';
        $this->title = '';
        $this->font = '';
        $this->name = '';
        $this->legendPos = 'r';
        $this->legendOverlay = 0;
        $this->border = '';
        $this->haxLabel = '';
        $this->vaxLabel = '';
        $this->haxLabelDisplay = '';
        $this->vaxLabelDisplay = '';
        $this->hgrid = '';
        $this->vgrid = '';
        $this->orientation = array();
        $this->axPos = array();

        $this->gapWidth = '';
        $this->overlap = '';
        $this->secondPieSize = '';
        $this->splitType = '';
        $this->splitPos = '';
        $this->custSplit = '';
        $this->subtype = '';

        $this->explosion = '';
        $this->holeSize = '';
        $this->symbol = '';
        $this->symbolSize = '';
        $this->style = '';
        $this->smooth = false;
        $this->wireframe = false;

        // default values for c:dLbls
        $this->showLegendKey = false;
        $this->showBubbleSize = false;
        $this->showPercent = false;
        $this->showValue = false;
        $this->showCategory = false;
        $this->showSeries = false;
        $this->showTable = false;

        $this->scalingMax = null;
        $this->scalingMin = null;

        $this->tickLblPos = 'nextTo';

        $this->majorUnit = null;
        $this->minorUnit = null;

        $this->stylesTitle = null;

        $this->formatDataLabels = null;
        $this->formatCode = null;
        $this->options = array();
    }

    /**
     * Setter. Rid
     *
     * @access public
     * @param string $rId
     */
    public function setRId($rId)
    {
        $this->rId = $rId;
    }

    /**
     * Getter. Rid
     *
     * @access public
     * @return string
     */
    public function getRId()
    {
        return $this->rId;
    }

    /**
     * Setter. Name
     *
     * @access public
     * @param string $name
     */
    public function setName($name)
    {
        $this->name = $name;
    }

    /**
     * Getter. Name
     *
     * @access public
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Create graphic
     *
     * @access public
     * @param array $chartOptions
     * @param array $options
     * @return string
     */
    public function createChart($chartOptions, $options)
    {
        $this->xmlChart = '';
        $this->setRId($options['rId']);
        $this->initGraphic($chartOptions, $options);
        $this->createEmbeddedXmlChart();

        return $this->xmlChart;
    }

    /**
     * Create embedded xml chart. To be replaced by chart type classes
     *
     * @access public
     */
    public function createEmbeddedXmlChart() {}

    /**
     * Init graphic
     *
     * @access public
     * @param array $options
     */
    public function initGraphic($chartOptions, $options)
    {
        $this->values = $chartOptions['data'];
        $this->options = $options;

        // chart options
        if (isset($chartOptions['theme'])) {
            $this->theme = $chartOptions['theme'];
        }
        if (isset($chartOptions['horizontalOffset'])) {
            $this->horizontalOffset = $chartOptions['horizontalOffset'];
        }
        if (isset($chartOptions['verticalOffset'])) {
            $this->verticalOffset = $chartOptions['verticalOffset'];
        }
        if (isset($chartOptions['showCategory']) && $chartOptions['showCategory']) {
            $this->showCategory = true;
        }
        if (isset($chartOptions['showLegendKey']) && $chartOptions['showLegendKey']) {
            $this->showLegendKey = true;
        }
        if (isset($chartOptions['showPercent']) && $chartOptions['showPercent']) {
            $this->showPercent = true;
        }
        if (isset($chartOptions['showSeries']) && $chartOptions['showSeries']) {
            $this->showSeries = true;
        }
        if (isset($chartOptions['showValue']) && $chartOptions['showValue']) {
            $this->showValue = true;
        }
        if (isset($chartOptions['rotX'])) {
            $this->rotX = $chartOptions['rotX'];
        }
        if (isset($chartOptions['rotY'])) {
            $this->rotY = $chartOptions['rotY'];
        }
        if (isset($chartOptions['perspective'])) {
            $this->perspective = $chartOptions['perspective'];
        }
        if (isset($chartOptions['color'])) {
            $this->color = $chartOptions['color'];
        }
        if (isset($chartOptions['groupBar'])) {
            $this->groupBar = $chartOptions['groupBar'];
        }
        if (isset($chartOptions['title'])) {
            $this->title = $chartOptions['title'];
        }
        if (isset($chartOptions['font'])) {
            $this->font = $chartOptions['font'];
        }
        if (isset($chartOptions['legendPos'])) {
            $this->legendPos = $chartOptions['legendPos'];
        }
        if (isset($chartOptions['legendOverlay']) && !empty($chartOptions['legendOverlay'])) {
            $this->legendOverlay = 1;
        }
        if (isset($chartOptions['border'])) {
            $this->border = $chartOptions['border'];
        }
        if (isset($chartOptions['haxLabel'])) {
            $this->haxLabel = $chartOptions['haxLabel'];
        }
        if (isset($chartOptions['vaxLabel'])) {
            $this->vaxLabel = $chartOptions['vaxLabel'];
        }
        if (isset($chartOptions['haxLabelDisplay'])) {
            $this->haxLabelDisplay = $chartOptions['haxLabelDisplay'];
        }
        if (isset($chartOptions['vaxLabelDisplay'])) {
            $this->vaxLabelDisplay = $chartOptions['vaxLabelDisplay'];
        }
        if (isset($chartOptions['showTable'])) {
            $this->showTable = $chartOptions['showTable'];
        }
        if (isset($chartOptions['hgrid'])) {
            $this->hgrid = $chartOptions['hgrid'];
        }
        if (isset($chartOptions['vgrid'])) {
            $this->vgrid = $chartOptions['vgrid'];
        }
        if (isset($chartOptions['style'])) {
            $this->style = $chartOptions['style'];
        }
        if (isset($chartOptions['gapWidth'])) {
            $this->gapWidth = $chartOptions['gapWidth'];
        }
        if (isset($chartOptions['overlap'])) {
            $this->overlap = $chartOptions['overlap'];
        }
        if (isset($chartOptions['secondPieSize'])) {
            $this->secondPieSize = $chartOptions['secondPieSize'];
        }
        if (isset($chartOptions['splitType'])) {
            $this->splitType = $chartOptions['splitType'];
        }
        if (isset($chartOptions['splitPos'])) {
            $this->splitPos = $chartOptions['splitPos'];
        }
        if (isset($chartOptions['custSplit'])) {
            $this->custSplit = $chartOptions['custSplit'];
        }
        if (isset($chartOptions['subtype'])) {
            $this->subtype = $chartOptions['subtype'];
        }
        if (isset($chartOptions['explosion'])) {
            $this->explosion = $chartOptions['explosion'];
        }
        if (isset($chartOptions['holeSize'])) {
            $this->holeSize = $chartOptions['holeSize'];
        }
        if (isset($chartOptions['majorUnit'])) {
            $this->majorUnit = $chartOptions['majorUnit'];
        }
        if (isset($chartOptions['minorUnit'])) {
            $this->minorUnit = $chartOptions['minorUnit'];
        }
        if (isset($chartOptions['scalingMax'])) {
            $this->scalingMax = $chartOptions['scalingMax'];
        }
        if (isset($chartOptions['scalingMin'])) {
            $this->scalingMin = $chartOptions['scalingMin'];
        }
        if (isset($chartOptions['stylesTitle'])) {
            $this->stylesTitle = $chartOptions['stylesTitle'];
        }
        if (isset($chartOptions['symbol'])) {
            $this->symbol = $chartOptions['symbol'];
        }
        if (isset($chartOptions['symbolSize'])) {
            $this->symbolSize = $chartOptions['symbolSize'];
        }
        if (isset($chartOptions['smooth'])) {
            $this->smooth = $chartOptions['smooth'];
        }
        if (isset($chartOptions['tickLblPos'])) {
            $this->tickLblPos = $chartOptions['tickLblPos'];
        }
        if (isset($chartOptions['wireframe'])) {
            $this->wireframe = $chartOptions['wireframe'];
        }
        if (isset($chartOptions['formatDataLabels'])) {
            $this->formatDataLabels = $chartOptions['formatDataLabels'];
        }
        if (isset($chartOptions['formatCode'])) {
            $this->formatCode = $chartOptions['formatCode'];
        }
        if (isset($chartOptions['orientation'])) {
            $this->orientation = $chartOptions['orientation'];
        }
        if (isset($chartOptions['axPos'])) {
            $this->axPos = $chartOptions['axPos'];
        }

        // extra options
        if (isset($options['type'])) {
            $this->type = $options['type'];
        }
    }

    /**
     * return the transposed matrix
     *
     * @access public
     * @param array
     */
    public function transposed($matrix)
    {
        $data = array();
        foreach ($matrix as $key => $value) {
            foreach ($value as $key2 => $value2) {
                $data[$key2][$key] = $value2;
            }
        }
        return $data;
    }

    /**
     * return the array with just 1 deep
     *
     * @access public
     * @param array
     */
    public function linear($matrix)
    {
        $data = array();
        foreach ($matrix as $key => $value) {
            foreach ($value as $ind => $val) {
                $data[] = $val;
            }
        }
        return $data;
    }

    /**
     * return the array of data prepared to modify the chart data
     *
     * @access public
     * @param array
     */
    public function prepareData($data)
    {
        $newData = array();
        $simple = true;
        if (isset($data['legend'])) {
            unset($data['legend']);
        }
        foreach ($data as $dat) {
            if (count($dat) > 1) {
                $simple = false;
            }
            break;
        }
        foreach ($data as $dat) {
            if ($simple) {
                $newData[] = $dat[0];
            } else {
                $newData[] = $dat;
            }
        }
        if ($simple) {
            return $this->linear(array($newData));
        } else {
            return $this->linear($this->transposed($newData));
        }
    }

    /**
     * Generate w:autotitledeleted
     *
     * @access protected
     * @param string $val
     */
    protected function generateAUTOTITLEDELETED($val = '1')
    {
        $xml = '<c:autoTitleDeleted val="' . $val . '"></c:autoTitleDeleted>__PHX=__GENERATECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:bar3DChart
     *
     * @access protected
     */
    protected function generateBAR3DCHART()
    {
        $xml = '<c:bar3DChart>__PHX=__GENERATETYPECHART__</c:bar3DChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:barChart
     *
     * @access protected
     */
    protected function generateBARCHART()
    {
        $xml = '<c:barChart>__PHX=__GENERATETYPECHART__</c:barChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:barDir
     *
     * @access protected
     * @param string $val
     */
    protected function generateBARDIR($val = 'bar')
    {
        $xml = '<c:barDir val="' . $val . '"></c:barDir>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:bodypr
     *
     * @access protected
     */
    protected function generateBODYPR()
    {
        $xml = '<a:bodyPr></a:bodyPr>__PHX=__GENERATERICH__';
        $this->xmlChart = str_replace('__PHX=__GENERATERICH__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:chart
     *
     * @access protected
     */
    protected function generateCHART()
    {
        $xml = '<c:chart>__PHX=__GENERATECHART__</c:chart>__PHX=__GENERATECHARTSPACE__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHARTSPACE__', $xml, $this->xmlChart);
    }

    /**
     * Generate chartspace XML
     *
     * @access protected
     */
    protected function generateCHARTSPACE()
    {
        $this->xmlChart = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">__PHX=__GENERATECHARTSPACE__</c:chartSpace>';
    }

    /**
     * Generate chartspace XML
     *
     * @access protected
     * @param array $serDataLabels
     */
    protected function generateDATALABELS_SER($serDataLabels, $idx) {
        // default values
        $position = 'ctr';
        $showCatName = 0;
        $showLegendKey = 0;
        $showPercent = 0;
        $showSerName = 0;
        $showVal = 0;

        $xml = '<c:dLbls>';
        if (isset($serDataLabels['formatCode'])) {
            $xml .= '<c:numFmt formatCode="'.$serDataLabels['formatCode'].'" sourceLinked="0"/>';
        }
        if (isset($serDataLabels['position'])) {
            switch ($serDataLabels['position']) {
                case 'center':
                    $position = 'ctr';
                    break;
                case 'insideBase':
                    $position = 'inBase';
                    break;
                case 'insideEnd':
                    $position = 'inEnd';
                    break;
                case 'outsideEnd':
                    $position = 'outEnd';
                    break;
                default:
                    $position = 'ctr';
                    break;
            }
        }

        $xml .= '<c:dLblPos val="'.$position.'"/>';
        if (isset($serDataLabels['showCategory'])) {
            $showCatName = $serDataLabels['showCategory'];
        }
        if (isset($serDataLabels['showLegendKey'])) {
            $showLegendKey = $serDataLabels['showLegendKey'];
        }
        if (isset($serDataLabels['showPercent'])) {
            $showPercent = $serDataLabels['showPercent'];
        }
        if (isset($serDataLabels['showSeries'])) {
            $showSerName = $serDataLabels['showSeries'];
        }
        if (isset($serDataLabels['showValue'])) {
            $showVal = $serDataLabels['showValue'];
        }
        $xml .= '<c:showLegendKey val="'.$showLegendKey.'"/><c:showVal val="'.$showVal.'"/><c:showCatName val="'.$showCatName.'"/><c:showSerName val="'.$showSerName.'"/><c:showPercent val="'.$showPercent.'"/><c:showBubbleSize val="0"/></c:dLbls>__PHX=__GENERATESER__';

        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:date1904
     *
     * @access protected
     * @param string $val
     */
    protected function generateDATE1904($val = '1')
    {
        $xml = '<c:date1904 val="' . $val . '"></c:date1904>__PHX=__GENERATECHARTSPACE__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHARTSPACE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:defrpr
     *
     * @access protected
     * @param string $styles style scope: title
     */
    protected function generateDEFRPR($scope = null)
    {
        if ($scope !== null && $this->stylesTitle !== null && is_array($this->stylesTitle)) {
            $stylesXML = '';
            $stylesExtraTagsXML = '';
            $stylesColorXML = '';

            if ($scope == 'title') {
                if (isset($this->stylesTitle['bold']) && $this->stylesTitle['bold']) {
                    $stylesXML .= ' b="1"';
                } else {
                    $stylesXML .= ' b="0"';
                }
                if (isset($this->stylesTitle['fontSize'])) {
                    $stylesXML .= ' sz="'.$this->stylesTitle['fontSize'].'"';
                } else {
                    $stylesXML .= ' sz="1420"';
                }
                if (isset($this->stylesTitle['italic']) && $this->stylesTitle['italic']) {
                    $stylesXML .= ' i="1"';
                } else {
                    $stylesXML .= ' i="0"';
                }

                if (isset($this->stylesTitle['color'])) {
                    $stylesColorXML .= '<a:solidFill><a:srgbClr val="'.$this->stylesTitle['color'].'"/></a:solidFill>';
                }

                if (isset($this->stylesTitle['font'])) {
                    $stylesColorXML .= '<a:latin typeface="'.$this->stylesTitle['font'].'"/>
                            <a:ea typeface="'.$this->stylesTitle['font'].'"/>
                            <a:cs typeface="'.$this->stylesTitle['font'].'"/>';
                }
            }

            $xml = '<a:defRPr'.$stylesXML.'>'.$stylesColorXML.'__PHX=__GENERATEDEFRPR__</a:defRPr>__PHX=__GENERATETITLEPPR__';
        } else {
            $xml = '<a:defRPr>__PHX=__GENERATEDEFRPR__</a:defRPr>__PHX=__GENERATETITLEPPR__';
        }
        $this->xmlChart = str_replace('__PHX=__GENERATETITLEPPR__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:lang
     *
     * @access protected
     * @param string $val
     */
    protected function generateLANG($val = 'en-US')
    {
        $phpxlsxconfig = \Phpxlsx\Utilities\PhpxlsxUtilities::parseConfig();
        if (isset($phpxlsxconfig['language'])) {
            $val = $phpxlsxconfig['language'];
        }
        $xml = '<c:lang val="' . $val . '"></c:lang>__PHX=__GENERATECHARTSPACE__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHARTSPACE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:line3DChart
     *
     * @access protected
     */
    protected function generateLINE3DCHART()
    {
        $xml = '<c:line3DChart>__PHX=__GENERATETYPECHART__</c:line3DChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:lineChart
     *
     * @access protected
     */
    protected function generateLINECHART()
    {
        $xml = '<c:lineChart>__PHX=__GENERATETYPECHART__</c:lineChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:area3DChart
     *
     * @access protected
     */
    protected function generateAREA3DCHART()
    {
        $xml = '<c:area3DChart>__PHX=__GENERATETYPECHART__</c:area3DChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:areaChart
     *
     * @access protected
     */
    protected function generateAREACHART()
    {
        $xml = '<c:areaChart>__PHX=__GENERATETYPECHART__</c:areaChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:perspective
     *
     * @access protected
     * @param string $val
     */
    protected function generatePERSPECTIVE($val = '30')
    {
        $xml = '<c:perspective val="' . $val . '"></c:perspective>';
        $this->xmlChart = str_replace('__PHX=__GENERATEVIEW3D__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:pie3DChart
     *
     * @access protected
     */
    protected function generatePIE3DCHART()
    {
        $xml = '<c:pie3DChart>__PHX=__GENERATETYPECHART__</c:pie3DChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:piechart
     *
     * @access protected
     */
    protected function generatePIECHART()
    {
        $xml = '<c:pieChart>__PHX=__GENERATETYPECHART__</c:pieChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:surfaceChart
     *
     * @access protected
     */
    protected function generateSURFACECHART()
    {
        $xml = '<c:surfaceChart>__PHX=__GENERATETYPECHART__</c:surfaceChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:wireframe
     *
     * @access protected
     */
    protected function generateWIREFRAME($val = 1)
    {
        $xml = '<c:wireframe val="' . $val . '" />__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:bubbleChart
     *
     * @access protected
     */
    protected function generateBUBBLECHART()
    {
        $xml = '<c:bubbleChart>__PHX=__GENERATETYPECHART__</c:bubbleChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:plotarea
     *
     * @access protected
     */
    protected function generatePLOTAREA()
    {
        $xml = '<c:plotArea>__PHX=__GENERATEPLOTAREA__</c:plotArea>__PHX=__GENERATECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:radarChart
     *
     * @access protected
     */
    protected function generateRADARCHART()
    {
        $xml = '<c:radarChart>__PHX=__GENERATETYPECHART__</c:radarChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:radarChart
     *
     * @access protected
     */
    protected function generateRADARCHARTSTYLE($style = 'radar')
    {
        $xml = '<c:radarStyle val="' . $style . '" />__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:rich
     *
     * @access protected
     */
    protected function generateRICH()
    {
        $xml = '<c:rich>__PHX=__GENERATERICH__</c:rich>__PHX=__GENERATETITLETX__';
        $this->xmlChart = str_replace('__PHX=__GENERATETITLETX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:rotx
     *
     * @access protected
     * @param string $val
     */
    protected function generateROTX($val = '30')
    {
        $xml = '<c:rotX val="'. $val . '"></c:rotX>__PHX=__GENERATEVIEW3D__';
        $this->xmlChart = str_replace('__PHX=__GENERATEVIEW3D__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:roty
     *
     * @access protected
     * @param string $val
     */
    protected function generateROTY($val = '30')
    {
        $xml = '<c:rotY val="' . $val . '"></c:rotY>__PHX=__GENERATEVIEW3D__';
        $this->xmlChart = str_replace('__PHX=__GENERATEVIEW3D__', $xml, $this->xmlChart);
    }

    /**
     * Generate rAngAx
     *
     * @access protected
     * @param string $val
     */
    protected function generateRANGAX($val = 0)
    {
        $xml = '<c:rAngAx val="' . $val . '"></c:rAngAx>__PHX=__GENERATEVIEW3D__';
        $this->xmlChart = str_replace('__PHX=__GENERATEVIEW3D__', $xml, $this->xmlChart);
    }

    /**
     * Generate roundedCorners
     *
     * @access protected
     * @param string $val
     */
    protected function generateROUNDEDCORNERS($val = 0)
    {
        $xml = '<c:roundedCorners val="' . $val . '"></c:roundedCorners>__PHX=__GENERATECHARTSPACE__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHARTSPACE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:style
     *
     * @access protected
     * @param string $val
     */
    protected function generateSTYLE($val = '2')
    {
        $style_2010 = (int) $val + 100;
        $xml = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14"><c14:style val="' . $style_2010 . '"/></mc:Choice><mc:Fallback><c:style val="' . $val . '"/></mc:Fallback></mc:AlternateContent>__PHX=__GENERATECHARTSPACE__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHARTSPACE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:title
     *
     * @access protected
     */
    protected function generateTITLE()
    {
        $xml = '<c:title>__PHX=__GENERATETITLE__<c:overlay val="0" /></c:title>__PHX=__GENERATECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:layout
     *
     * @access protected
     */
    protected function generateLAYOUT()
    {
        $xml = '<c:layout></c:layout>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:titlelayout
     *
     * @access protected
     * @param string $nombre
     */
    protected function generateTITLELAYOUT($nombre = '')
    {
        $xml = '<a:layout></a:layout>';
        $this->xmlChart = str_replace('__PHX=__GENERATETITLE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:titlep
     *
     * @access protected
     */
    protected function generateTITLEP()
    {
        $xml = '<a:p>__PHX=__GENERATETITLEP__</a:p>__PHX=__GENERATERICH__';
        $this->xmlChart = str_replace('__PHX=__GENERATERICH__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:titleppr
     *
     * @access protected
     */
    protected function generateTITLEPPR()
    {
        $xml = '<a:pPr>__PHX=__GENERATETITLEPPR__</a:pPr>__PHX=__GENERATETITLEP__';
        $this->xmlChart = str_replace('__PHX=__GENERATETITLEP__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:titler
     *
     * @access protected
     */
    protected function generateTITLER()
    {
        $xml = '<a:r>__PHX=__GENERATETITLER__</a:r>__PHX=__GENERATETITLEP__';
        $this->xmlChart = str_replace('__PHX=__GENERATETITLEP__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:titlerfonts
     *
     * @access protected
     * @param string $font
     */
    protected function generateTITLERFONTS($font = '')
    {
        $xml = '<a:latin typeface="' . $font . '" pitchFamily="34" charset="0"></a:latin ><a:cs typeface="' . $font . '" pitchFamily="34" charset="0"></a:cs>';
        $this->xmlChart = str_replace('__PHX=__GENERATETITLERPR__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:titlerpr
     *
     * @access protected
     */
    protected function generateTITLERPR($lang = 'es-ES')
    {
        $xml = '<a:rPr lang="' . $lang . '">__PHX=__GENERATETITLERPR__</a:rPr>__PHX=__GENERATETITLER__';
        $this->xmlChart = str_replace('__PHX=__GENERATETITLER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:titlet
     *
     * @access protected
     * @param string $title
     */
    protected function generateTITLET($title = '')
    {
        $xml = '<a:t>' . htmlspecialchars($title) . '</a:t>__PHX=__GENERATETITLER__';
        $this->xmlChart = str_replace('__PHX=__GENERATETITLER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:titletx
     *
     * @access protected
     */
    protected function generateTITLETX()
    {
        $xml = '<c:tx>__PHX=__GENERATETITLETX__</c:tx>__PHX=__GENERATETITLE__';
        $this->xmlChart = str_replace('__PHX=__GENERATETITLE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:varyColors
     *
     * @access protected
     * @param string $val
     */
    protected function generateVARYCOLORS($val = '1')
    {
        $xml = '<c:varyColors val="' . $val . '"></c:varyColors>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:view3D
     *
     * @access protected
     */
    protected function generateVIEW3D()
    {
        $xml = '<c:view3D>__PHX=__GENERATEVIEW3D__</c:view3D>__PHX=__GENERATECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:grouping
     *
     * @access protected
     * @param string $val
     */
    protected function generateGROUPING($val = 'stacked')
    {
        $xml = '<c:grouping val="' . $val . '"></c:grouping>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:ser
     *
     * @access protected
     */
    protected function generateSER()
    {
        $xml = '<c:ser>__PHX=__GENERATESER__</c:ser>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:idx
     *
     * @access protected
     * @param string $val
     */
    protected function generateIDX($val = '0')
    {
        $xml = '<c:idx val="' . $val . '"></c:idx>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:order
     *
     * @access protected
     * @param string $val
     */
    protected function generateORDER($val = '0')
    {
        $xml = '<c:order val="' . $val . '"></c:order>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:tx
     *
     * @access protected
     */
    protected function generateTX()
    {
        $xml = '<c:tx>__PHX=__GENERATETX__</c:tx>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:dLbls
     *
     * @access protected
     */
    protected function generateSERDLBLS()
    {
        $xml = '<c:dLbls>__PHX=__GENERATEDLBLS__</c:dLbls>__PHX=__GENERATETYPECHART__';

        if ($this->formatDataLabels !== null) {
            $rotation = 0;
            if(isset($this->formatDataLabels['rotation'])) {
                $rotation = 60000 * $this->formatDataLabels['rotation'];
            }
            $position = 'outEnd';
            if (isset($this->formatDataLabels['position'])) {
                switch ($this->formatDataLabels['position']) {
                    case 'center':
                        $position = 'ctr';
                        break;
                    case 'insideBase':
                        $position = 'inBase';
                        break;
                    case 'insideEnd':
                        $position = 'inEnd';
                        break;
                    case 'outsideEnd':
                        $position = 'outEnd';
                        break;
                    default:
                        $position = 'outEnd';
                        break;
                }
            }

            $xmlFormatDataLabels = '<c:txPr><a:bodyPr anchor="ctr" anchorCtr="1" bIns="19050" lIns="38100" rIns="38100" rot="'.$rotation.'" spcFirstLastPara="1" tIns="19050" vertOverflow="ellipsis" wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr b="0" baseline="0" i="0" kern="1200" strike="noStrike" sz="900" u="none"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="es-ES"/></a:p></c:txPr><c:dLblPos val="'.$position.'"/>__PHX=__GENERATEDLBLS__';

            $xml = str_replace('__PHX=__GENERATEDLBLS__', $xmlFormatDataLabels, $xml);
        }

        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:showBubbleSize
     *
     * @access protected
     */
    protected function generateSHOWBUBBLESIZE($val = true)
    {
        $value = $val ? '1' : '0';
        $xml = '<c:showBubbleSize val="' . $value . '"></c:showBubbleSize>';
        $this->xmlChart = str_replace('__PHX=__GENERATEDLBLS__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:showLegendKey
     *
     * @access protected
     */
    protected function generateSHOWLEGENDKEY($val = true)
    {
        $value = $val ? '1' : '0';
        $xml = '<c:showLegendKey val="' . $value . '"></c:showLegendKey>__PHX=__GENERATEDLBLS__';
        $this->xmlChart = str_replace('__PHX=__GENERATEDLBLS__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:showVal
     *
     * @access protected
     */
    protected function generateSHOWVAL($val = 1)
    {
        $value = $val ? '1' : '0';
        $xml = '<c:showVal val="' . $value . '"></c:showVal>__PHX=__GENERATEDLBLS__';
        $this->xmlChart = str_replace('__PHX=__GENERATEDLBLS__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:showCatName
     *
     * @access protected
     */
    protected function generateSHOWCATNAME($val = true)
    {
        $value = $val ? '1' : '0';
        $xml = '<c:showCatName val="' . $value . '"></c:showCatName>__PHX=__GENERATEDLBLS__';
        $this->xmlChart = str_replace('__PHX=__GENERATEDLBLS__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:showSerName
     *
     * @access protected
     */
    protected function generateSHOWSERNAME($val = true)
    {
        $value = $val ? '1' : '0';
        $xml = '<c:showSerName val="' . $value . '"></c:showSerName>__PHX=__GENERATEDLBLS__';
        $this->xmlChart = str_replace('__PHX=__GENERATEDLBLS__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:strref
     *
     * @access protected
     */
    protected function generateSTRREF()
    {
        $xml = '<c:strRef>__PHX=__GENERATESTRREF__</c:strRef>__PHX=__GENERATETX__';
        $this->xmlChart = str_replace('__PHX=__GENERATETX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:f
     *
     * @access protected
     * @param string $val
     */
    protected function generateF($val = 'Sheet1!$B$1')
    {
        $xml = '<c:f>' . $val . '</c:f>__PHX=__GENERATESTRREF__';
        $this->xmlChart = str_replace('__PHX=__GENERATESTRREF__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:strcache
     *
     * @access protected
     */
    protected function generateSTRCACHE()
    {
        $xml = '<c:strCache>__PHX=__GENERATESTRCACHE__</c:strCache>__PHX=__GENERATESTRREF__';
        $this->xmlChart = str_replace('__PHX=__GENERATESTRREF__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:ptcount
     *
     * @access protected
     * @param string $val
     */
    protected function generatePTCOUNT($val = '1')
    {
        $xml = '<c:ptCount val="' . $val . '"></c:ptCount>__PHX=__GENERATESTRCACHE__';
        $this->xmlChart = str_replace('__PHX=__GENERATESTRCACHE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:pt
     *
     * @access protected
     * @param string $idx
     */
    protected function generatePT($idx = '0')
    {
        $xml = '<c:pt idx="' . $idx . '">__PHX=__GENERATEPT__</c:pt>__PHX=__GENERATESTRCACHE__';
        $this->xmlChart = str_replace('__PHX=__GENERATESTRCACHE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:v
     *
     * @access protected
     * @param string $idx
     */
    protected function generateV($idx = 'Ventas')
    {
        $xml = '<c:v>' . htmlspecialchars($idx) . '</c:v>';
        $this->xmlChart = str_replace('__PHX=__GENERATEPT__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:trendline
     *
     * @access protected
     */
    protected function generateTRENDLINE($trendline)
    {
        if (count($trendline) > 0) {
            $xmlTRENDLINE = '<c:trendline>';

            if (isset($trendline['color']) || isset($trendline['lineStyle'])) {
                $xmlTRENDLINE .= '<c:spPr><a:ln cap="rnd" w="19050">';
                if (isset($trendline['color'])) {
                    $xmlTRENDLINE .= '<a:solidFill><a:srgbClr val="'.$trendline['color'].'"/></a:solidFill>';
                }
                if (isset($trendline['lineStyle'])) {
                    $xmlTRENDLINE .= '<a:prstDash val="'.$trendline['lineStyle'].'"/>';
                }
                $xmlTRENDLINE .= '</a:ln><a:effectLst/></c:spPr>';
            }

            if (!isset($trendline['type'])) {
                $trendline['type'] = 'linear';
            }
            $xmlTRENDLINE .= '<c:trendlineType val="'.$trendline['type'].'"/>';

            if (isset($trendline['typeOrder'])) {
                if ($trendline['type'] == 'poly') {
                    $xmlTRENDLINE .= '<c:order val="'.$trendline['typeOrder'].'"/>';
                }
                if ($trendline['type'] == 'movingAvg') {
                    $xmlTRENDLINE .= '<c:period val="'.$trendline['typeOrder'].'"/>';
                }
            } else {
                if ($trendline['type'] == 'poly') {
                    $xmlTRENDLINE .= '<c:order val="2"/>';
                }
                if ($trendline['type'] == 'movingAvg') {
                    $xmlTRENDLINE .= '<c:period val="2"/>';
                }
            }

            if (isset($trendline['intercept'])) {
                $xmlTRENDLINE .= '<c:intercept val="'.$trendline['intercept'].'"/>';
            }
            if (isset($trendline['displayRSquared']) && $trendline['displayRSquared'] == true) {
                $xmlTRENDLINE .= '<c:dispRSqr val="1"/>';
            } else {
                $xmlTRENDLINE .= '<c:dispRSqr val="0"/>';
            }
            if (isset($trendline['displayEquation']) && $trendline['displayEquation'] == true) {
                $xmlTRENDLINE .= '<c:dispEq val="1"/>';
            } else {
                $xmlTRENDLINE .= '<c:dispEq val="0"/>';
            }

            $xmlTRENDLINE .= '</c:trendline>';

            $xml = $xmlTRENDLINE.'__PHX=__GENERATESER__';

            $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
        }
    }

    /**
     * Generate w:cat
     *
     * @access protected
     */
    protected function generateCAT()
    {
        $xml = '<c:cat>__PHX=__GENERATETX__</c:cat>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:val
     *
     * @access protected
     */
    protected function generateVAL()
    {
        $xml = '<c:val>__PHX=__GENERATETX__</c:val>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:numcache
     *
     * @access protected
     */
    protected function generateNUMCACHE()
    {
        $xml = '<c:numCache>__PHX=__GENERATESTRCACHE__</c:numCache>__PHX=__GENERATESTRREF__';
        $this->xmlChart = str_replace('__PHX=__GENERATESTRREF__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:layout
     *
     * @access protected
     * @param string $font
     */
    protected function generateLEGENDLAYOUT()
    {
        $xml = '<c:layout />__PHX=__GENERATELEGEND__';
        $this->xmlChart = str_replace('__PHX=__GENERATELEGEND__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:xVal
     *
     * @access protected
     */
    protected function generateXVAL()
    {
        $xml = '<c:xVal>__PHX=__GENERATETX__</c:xVal>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:spPr
     *
     * @access protected
     */
    protected function generateSPPR_SER()
    {
        $xml = '<c:spPr>__PHX=__GENERATESPPR__</c:spPr>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate a:solidFill
     *
     * @access protected
     * @param string $color
     */
    protected function generateSPPR_SOLIDFILL($color)
    {
        $xml = '<a:solidFill>
                    <a:srgbClr val="'.$color.'"/>
                </a:solidFill>
                <a:ln>
                    <a:noFill/>
                </a:ln>
                <a:effectLst/>__PHX=__GENERATESPPR__';

        $this->xmlChart = str_replace('__PHX=__GENERATESPPR__', $xml, $this->xmlChart);
    }

    /**
     * Generate a:cdpt
     *
     * @access protected
     * @param array $values
     */
    protected function generateCDPT($values)
    {
        $xml = '';
        for ($i = 0; $i < count($values); $i++) {
            if ($values[$i] == null) {
                continue;
            }
            $xml .= '<c:dPt><c:idx val="'.$i.'"/>
                    <c:spPr>
                        <a:solidFill>
                            <a:srgbClr val="'.$values[$i].'"/>
                        </a:solidFill>
                        <a:ln>
                            <a:noFill/>
                        </a:ln>
                        <a:effectLst/>
                    </c:spPr></c:dPt>
            ';
        }
        $xml .= '__PHX=__GENERATESER__';

        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:yVal
     *
     * @access protected
     */
    protected function generateYVAL()
    {
        $xml = '<c:yVal>__PHX=__GENERATETX__</c:yVal>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:bubbleSize
     *
     * @access protected
     */
    protected function generateBUBBLESIZE()
    {
        $xml = '<c:bubbleSize>__PHX=__GENERATETX__</c:bubbleSize>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:smooth
     *
     * @access protected
     */
    protected function generateSMOOTH($val = 1)
    {
        $xml = '<c:smooth val="' . $val . '">__PHX=__GENERATETX__</c:smooth>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:bubble3D
     *
     * @access protected
     */
    protected function generateBUBBLES3D($val = 1)
    {
        $xml = '<c:bubble3D val="' . $val . '"></c:bubble3D>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:bubbleScale
     *
     * @access protected
     */
    protected function generateBUBBLESCALE($val = 100)
    {
        $xml = '<c:bubbleScale val="' . $val . '"></c:bubbleScale>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate a:txPr
     *
     * @access protected
     * @param string $font
     */
    protected function generateTXPR()
    {
        $xml = '<c:txPr>__PHX=__GENERATETXPR__</c:txPr>__PHX=__GENERATECHARTSPACE__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHARTSPACE__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:bodyPr
     *
     * @access protected
     */
    protected function generateLEGENDBODYPR()
    {
        $xml = '<a:bodyPr></a:bodyPr>__PHX=__GENERATERICH__';
        $this->xmlChart = str_replace('__PHX=__GENERATETXPR__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:lststyle
     *
     * @access protected
     */
    protected function generateLSTSTYLE()
    {
        $xml = '<a:lstStyle></a:lstStyle>__PHX=__GENERATERICH__';
        $this->xmlChart = str_replace('__PHX=__GENERATERICH__', $xml, $this->xmlChart);
    }

    /**
     * Generate a:p
     *
     * @access protected
     */
    protected function generateAP()
    {
        $xml = '<a:p>__PHX=__GENERATEAP__</a:p>__PHX=__GENERATERICH__';
        $this->xmlChart = str_replace('__PHX=__GENERATERICH__', $xml, $this->xmlChart);
    }

    /**
     * Generate a:pPr
     *
     * @access protected
     */
    protected function generateAPPR($rtl = 0)
    {
        $xml = '<a:pPr rtl="' . $rtl . '">__PHX=__GENERATETITLEPPR__</a:pPr>__PHX=__GENERATEAP__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAP__', $xml, $this->xmlChart);
    }

    /**
     * Generate a:endParaRPr
     *
     * @access protected
     */
    protected function generateENDPARARPR($lang = "es-ES_tradnl")
    {
        $xml = '<a:endParaRPr lang="' . $lang . '" />__PHX=__GENERATEAP__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAP__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:numRef
     *
     * @access protected
     */
    protected function generateNUMREF()
    {
        $xml = '<c:numRef>__PHX=__GENERATESTRREF__</c:numRef>__PHX=__GENERATETX__';
        $this->xmlChart = str_replace('__PHX=__GENERATETX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:formatCode
     *
     * @access protected
     * @param string $val
     */
    protected function generateFORMATCODE($val = 'General')
    {
        $this->xmlChart = str_replace('__PHX=__GENERATESTRCACHE__', '<c:formatCode>' . $val .
                '</c:formatCode>__PHX=__GENERATESTRCACHE__', $this->xmlChart);
    }

    /**
     * Generate w:legend
     *
     * @access protected
     */
    protected function generateLEGEND()
    {
        if ($this->legendPos != 'none') {
            $xml = '<c:legend>__PHX=__GENERATELEGEND__</c:legend>__PHX=__GENERATECHART__';
            $this->xmlChart = str_replace('__PHX=__GENERATECHART__', $xml, $this->xmlChart);
        }
    }

    /**
     * Generate c:legendPos
     *
     * @access protected
     * @param string $val
     */
    protected function generateLEGENDPOS($val = 'r')
    {
        if ($val != 'none') {
            $xml = '<c:legendPos val="' . $val . '"></c:legendPos>__PHX=__GENERATELEGEND__';
            $this->xmlChart = str_replace('__PHX=__GENERATELEGEND__', $xml, $this->xmlChart);
        }
    }

    /**
     * Generate c:layout
     *
     * @access protected
     * @param string $font
     */
    protected function generateLEGENDFONT($font = '')
    {
        $xml = '<c:layout /><c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr><a:latin typeface="' . $font . '" pitchFamily="34" charset="0" /><a:cs typeface="' . $font . '" pitchFamily="34" charset="0" /></a:defRPr></a:pPr><a:endParaRPr lang="es-ES" /></a:p></c:txPr>';
        $this->xmlChart = str_replace('__PHX=__GENERATELEGEND__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:overlay
     *
     * @access protected
     * @param string $val
     */
    protected function generateLEGENDOVERLAY($val = 0)
    {
        $xml = '<c:overlay val="'. $val . '" />__PHX=__GENERATELEGEND__';
        $this->xmlChart = str_replace('__PHX=__GENERATELEGEND__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:poltVisOnly
     *
     * @access protected
     * @param string $val
     */
    protected function generatePLOTVISONLY($val = '1')
    {
        $xml = '<c:plotVisOnly val="'. $val . '"></c:plotVisOnly>__PHX=__GENERATECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:spPr
     *
     * @access protected
     */
    protected function generateSPPR()
    {
        $xml = '<c:spPr>__PHX=__GENERATESPPR__</c:spPr>__PHX=__GENERATECHARTSPACE__';
        $this->xmlChart = str_replace('__PHX=__GENERATECHARTSPACE__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:ln
     *
     * @access protected
     */
    protected function generateLN($w = NULL)
    {
        if (is_numeric($w)) {
            $xml = '<a:ln w="' . ($w * 12700) . '">__PHX=__GENERATELN__</a:ln>';
        } else {
            $xml = '<a:ln>__PHX=__GENERATELN__</a:ln>';
        }
        $this->xmlChart = str_replace('__PHX=__GENERATESPPR__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:noFill
     *
     * @access protected
     */
    protected function generateNOFILL()
    {
        $xml = '<a:noFill></a:noFill>';
        $this->xmlChart = str_replace('__PHX=__GENERATELN__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:overlap
     *
     * @access protected
     * @param string $val
     */
    protected function generateOVERLAP($val = '100')
    {
        $xml = '<c:overlap val="'. $val . '"></c:overlap>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:shape
     *
     * @access protected
     * @param string $val
     */
    protected function generateSHAPE($val = 'box')
    {
        $xml = '<c:shape val="'. $val . '"></c:shape>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:bandFmts
     *
     * @access protected
     * @param string $val
     */
    protected function generateBANDFMTS($val = 'box')
    {
        $xml = '<c:bandFmts />__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:axid
     *
     * @access protected
     * @param string $val
     */
    protected function generateAXID($val = '59034624')
    {
        $xml = '<c:axId val="'. $val . '"></c:axId>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:firstSliceAng
     *
     * @access protected
     * @param string $val
     */
    protected function generateFIRSTSLICEANG($val = '0')
    {
        $xml = '<c:firstSliceAng val="' . $val . '"></c:firstSliceAng>' . '__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:dLbls
     *
     * @access protected
     */
    protected function generateDLBLS()
    {
        $xml = '<c:dLbls>__PHX=__GENERATEDLBLS__</c:dLbls>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:holeSize
     *
     * @access protected
     */
    protected function generateHOLESIZE($val = 50)
    {
        $xml = '<c:holeSize val="' . $val . '"></c:holeSize>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:showPercent
     *
     * @access protected
     * @param string $val
     */
    protected function generateSHOWPERCENT($val = false)
    {
        $value = $val ? '1' : '0';
        $xml = '<c:showPercent val="' . $value . '"></c:showPercent>__PHX=__GENERATEDLBLS__';
        $this->xmlChart = str_replace('__PHX=__GENERATEDLBLS__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:catAx
     *
     * @access protected
     */
    protected function generateCATAX()
    {
        $xml = '<c:catAx>__PHX=__GENERATEAX__</c:catAx>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:dTable
     *
     * @access protected
     */
    protected function generateDATATABLE()
    {
        $xml = '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:valAx
     *
     * @access protected
     */
    protected function generateVALAX()
    {
        $xml = '<c:valAx>__PHX=__GENERATEAX__</c:valAx>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:axId
     *
     * @access protected
     * @param <type> $val
     */
    protected function generateAXAXID($val = '59034624')
    {
        $xml = '<c:axId val="'. $val . '"></c:axId>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:scaling
     *
     * @access protected
     */
    protected function generateDELETE($val = 0)
    {
        $xml = '<c:delete val="' . $val . '"></c:delete>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:scaling
     *
     * @access protected
     * @param bool $addScalingValues
     */
    protected function generateSCALING($addScalingValues = false)
    {
        $xml = '<c:scaling>__PHX=__GENERATESCALING__</c:scaling>__PHX=__GENERATEAX__';

        if ($this->scalingMax !== null && $addScalingValues) {
            $xml = str_replace('__PHX=__GENERATESCALING__', '<c:max val="'.$this->scalingMax.'" />__PHX=__GENERATESCALING__', $xml);
        }

        if ($this->scalingMin !== null && $addScalingValues) {
            $xml = str_replace('__PHX=__GENERATESCALING__', '<c:min val="'.$this->scalingMin.'" />__PHX=__GENERATESCALING__', $xml);
        }

        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:orientation
     *
     * @access protected
     * @param string $val
     */
    protected function generateORIENTATION($val = 'minMax')
    {
        $xml = '<c:orientation val="'. $val . '"></c:orientation>';
        $this->xmlChart = str_replace('__PHX=__GENERATESCALING__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:axPos
     *
     * @access protected
     * @param string $val
     */
    protected function generateAXPOS($val = 'b')
    {
        $xml = '<c:axPos val="' . $val . '"></c:axPos>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:title
     *
     * @access protected
     * @param string $val
     */
    protected function generateAXLABEL($val = 'Axis title')
    {
        $xml = '<c:title><c:tx><c:rich>__PHX=__GENERATEBODYPR__<a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>' . htmlspecialchars($val) . '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate a:bodyPr
     *
     * @access protected
     * @param string $val
     */
    protected function generateAXLABELDISP($val = 'horz', $rot = 0)
    {
        $xml = '<a:bodyPr rot="' . $rot . '" vert="' . $val . '"/>';
        $this->xmlChart = str_replace('__PHX=__GENERATEBODYPR__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:surface3DChart
     *
     * @access protected
     */
    protected function generateSURFACE3DCHART()
    {
        $xml = '<c:surface3DChart>__PHX=__GENERATETYPECHART__</c:surface3DChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:serAx
     *
     * @access protected
     */
    protected function generateSERAX()
    {
        $xml = '<c:serAx>__PHX=__GENERATEAX__</c:serAx>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:scatterStyle
     *
     * @access protected
     */
    protected function generateSCATTERSTYLE($style = 'smoothMarker')
    {
        $possibleStyles = array('none', 'line', 'lineMarker', 'marker', 'smooth', 'smoothMarker');
        if (!in_array($style, $possibleStyles)) {
            $style = 'smoothMarker';
        }
        $xml = '<c:scatterStyle val="' . $style . '" />__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:tickLblPos
     *
     * @access protected
     * @param string $val
     * @param bool $isHorizontal
     */
    protected function generateTICKLBLPOS($val = 'nextTo', $isHorizontal = false)
    {
        if ($isHorizontal) {
            $val = $this->tickLblPos;
        }

        $xml = '<c:tickLblPos val="'. $val . '"></c:tickLblPos>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:crossAx
     *
     * @access protected
     * @param string $val
     */
    protected function generateCROSSAX($val = '59040512')
    {
        $xml = '<c:crossAx  val="'. $val . '"></c:crossAx >__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:crosses
     *
     * @access protected
     * @param string $val
     */
    protected function generateCROSSES($val = 'autoZero')
    {
        $xml = '<c:crosses val="'. $val . '"></c:crosses>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:auto
     *
     * @access protected
     * @param string $val
     */
    protected function generateAUTO($val = '1')
    {
        $xml = '<c:auto val="'. $val . '"></c:auto>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:lblAlgn
     *
     * @access protected
     * @param string $val
     */
    protected function generateLBLALGN($val = 'ctr')
    {
        $xml = '<c:lblAlgn val="'. $val . '"></c:lblAlgn>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:lblOffset
     *
     * @access protected
     * @param string $val
     */
    protected function generateLBLOFFSET($val = '100')
    {
        $xml = '<c:lblOffset val="'. $val . '"></c:lblOffset>';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:majorTickMark
     *
     * @access protected
     */
    protected function generateMAJORTICKMARK($val = 'none')
    {
        $xml = '<c:majorTickMark val="' . $val . '"></c:majorTickMark>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:majorUnit
     *
     * @access protected
     */
    protected function generateMAJORUNIT($val = null)
    {
        if ($val !== null) {
            $xml = '<c:majorUnit val="' . $val . '"></c:majorUnit>__PHX=__GENERATEAX__';
            $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
        }
    }

    /**
     * Generate c:minorUnit
     *
     * @access protected
     */
    protected function generateMINORUNIT($val = null)
    {
        if ($val !== null) {
            $xml = '<c:minorUnit val="' . $val . '"></c:minorUnit>__PHX=__GENERATEAX__';
            $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
        }
    }

    /**
     * Generate c:majorGridlines
     *
     * @access protected
     */
    protected function generateMAJORGRIDLINES()
    {
        $xml = '<c:majorGridlines></c:majorGridlines>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:majorGridlines
     *
     * @access protected
     */
    protected function generateMARKER($symbol = 'none', $size = NULL)
    {
        $symbols = array('circle', 'dash', 'diamond', 'dot', 'none', 'picture', 'plus', 'square', 'star', 'triangle', 'x');
        if (!in_array($symbol, $symbols)) {
            $symbol = 'none';
        }
        $xml = '<c:marker><c:symbol val="' . $symbol . '"/>';
        if (!empty($size) && is_int($size) && $size < 73 && $size > 1) {
            $xml .= '<c:size val="' . $size . '"></c:size>';
        }
        $xml .= '</c:marker>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:majorGridlines
     *
     * @access protected
     */
    protected function generateMINORGRIDLINES($val = '')
    {
        $xml = '<c:minorGridlines></c:minorGridlines>__PHX=__GENERATEAX__';
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:numFmt
     *
     * @access protected
     * @param string $formatCode
     * @param string $sourceLinked
     */
    protected function generateNUMFMT($formatCode = 'General', $sourceLinked = '1')
    {
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', '<c:numFmt formatCode="' . $formatCode . '" sourceLinked="' . $sourceLinked . '"></c:numFmt>__PHX=__GENERATEAX__', $this->xmlChart);
    }

    /**
     * Generate w:numFmt in ser
     *
     * @access protected
     */
    protected function generateNUMFMT_SER($formatCode = 'General', $sourceLinked = '1')
    {
        $xml = '<c:numFmt formatCode="' . $formatCode . '" sourceLinked="' . $sourceLinked . '" />__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:latin
     *
     * @access protected
     * @param string $font
     */
    protected function generateRFONTS($font)
    {
        $this->xmlChart = str_replace('__PHX=__GENERATEDEFRPR__', '<a:latin typeface="' . $font . '" pitchFamily="34" charset="0"></a:latin ><a:cs typeface="' . $font . '" pitchFamily="34" charset="0"></a:cs>', $this->xmlChart);
    }

    /**
     * Generate w:crossBetween
     *
     * @access protected
     * @param string $val
     */
    protected function generateCROSSBETWEEN($val = 'between')
    {
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', '<c:crossBetween val="'. $val . '"></c:crossBetween>', $this->xmlChart);
    }

    /**
     * Generate w:ofPieChart
     *
     * @access protected
     * @param string $val
     */
    protected function generateOFPIECHART()
    {
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', '<c:ofPieChart>__PHX=__GENERATETYPECHART__</c:ofPieChart>', $this->xmlChart);
    }

    /**
     * Generate c:ofPieType
     *
     * @access protected
     * @param string $val
     */
    protected function generateOFPIETYPE($val = 'pie')
    {
        if (!in_array($val, array('pie', 'bar'))) {
            $val = 'pie';
        }
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', '<c:ofPieType val="' . $val . '"></c:ofPieType>__PHX=__GENERATETYPECHART__', $this->xmlChart);
    }

    /**
     * Generate w:scatterChart
     *
     * @access protected
     */
    protected function generateSCATTERCHART()
    {
        $xml = '<c:scatterChart>__PHX=__GENERATETYPECHART__</c:scatterChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate w:doughnutChart
     *
     * @access protected
     */
    protected function generateDOUGHNUTCHART()
    {
        $xml = '<c:doughnutChart>__PHX=__GENERATETYPECHART__</c:doughnutChart>__PHX=__GENERATEPLOTAREA__';
        $this->xmlChart = str_replace('__PHX=__GENERATEPLOTAREA__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:GAPWIDTH
     *
     * @access protected
     * @param string $val
     */
    protected function generateGAPWIDTH($val = 100)
    {
        if (!is_numeric($val)) {
            $val = 100;
        }
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', '<c:gapWidth val="' . $val . '"></c:gapWidth>__PHX=__GENERATETYPECHART__', $this->xmlChart);
    }

    /**
     * Generate c:secondPieSize
     *
     * @access protected
     * @param string $val
     */
    protected function generateSECONDPIESIZE($val = 75)
    {
        if (!is_numeric($val)) {
            $val = 75;
        }
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', '<c:secondPieSize val="' . $val . '"></c:secondPieSize>__PHX=__GENERATETYPECHART__', $this->xmlChart);
    }

    /**
     * Generate c:serLines
     *
     * @access protected
     */
    protected function generateSERLINES()
    {
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', '<c:serLines></c:serLines>__PHX=__GENERATETYPECHART__', $this->xmlChart);
    }

    /**
     * Generate c:splitType
     *
     * @access protected
     * @param string $val
     */
    protected function generateSPLITTYPE($val)
    {
        if (!in_array($val, array('auto', 'cust', 'percent', 'pos', 'val'))) {
            $xml = '<c:splitType></c:splitType>__PHX=__GENERATETYPECHART__';
        } else {
            $xml = '<c:splitType val="' . $val . '"></c:splitType>__PHX=__GENERATETYPECHART__';
        }
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:custSplit
     *
     * @access protected
     */
    protected function generateCUSTSPLIT()
    {
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', '<c:custSplit>__PHX=__GENERATECUSTSPLIT__</c:custSplit>__PHX=__GENERATETYPECHART__', $this->xmlChart);
    }

    /**
     * Generate c:splitType
     *
     * @access protected
     * @param string $val
     */
    protected function generateSECONDPIEPT($val)
    {
        $xml = '';
        if (is_array($val)) {
            foreach ($val as $value) {
                $xml .= '<c:secondPiePt val="' . $value . '"></c:secondPiePt>';
            }
        }
        $this->xmlChart = str_replace('__PHX=__GENERATECUSTSPLIT__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:splitPos
     *
     * @access protected
     * @param string $val
     */
    protected function generateSPLITPOS($val, $type = "auto")
    {
        if ($type == 'pos') {
            $val = (int) $val;
        }
        $xml = '<c:splitPos val="' . $val . '"></c:splitPos>__PHX=__GENERATETYPECHART__';
        $this->xmlChart = str_replace('__PHX=__GENERATETYPECHART__', $xml, $this->xmlChart);
    }

    /**
     * Generate c:explosion
     *
     * @access protected
     * @param string $val
     */
    protected function generateEXPLOSION($val = 25)
    {
        $xml = '<c:explosion val="' . $val . '"></c:explosion>__PHX=__GENERATESER__';
        $this->xmlChart = str_replace('__PHX=__GENERATESER__', $xml, $this->xmlChart);
    }

    /**
     * Clean tags in template document
     *
     * @access protected
     */
    protected function cleanTemplateDocument()
    {
        $this->xmlChart = preg_replace('/__PHX=__[A-Z]+__/', '', $this->xmlChart);
    }

    /**
     * Clean tags in template document
     *
     * @access protected
     */
    public static function cleanTemplateChart($xml = "")
    {
        return preg_replace('/__PHX=__[A-Z]+__/', '', $xml);
    }

    /**
     * Clean tags in template document
     *
     * @access protected
     */
    protected function cleanTemplate2()
    {
        $this->xmlChart = preg_replace(
                array(
            '/__PHX=__GENERATE[A-B,D-O,Q-R,U-Z][A-Z]+__/',
            '/__PHX=__GENERATES[A-D,F-Z][A-Z]+__/', '/__PHX=__GENERATETX__/'), '', $this->xmlChart);
    }

    /**
     * Clean tags in template document
     *
     * @access protected
     */
    protected function cleanTemplateFonts()
    {
        $this->xmlChart = preg_replace(
                '/__PHX=__GENERATETITLE[A-Z]+__/', '', $this->xmlChart);
    }

    /**
     * Clean tags in template document
     *
     * @access protected
     */
    protected function cleanTemplate3()
    {
        $this->xmlChart = preg_replace(
                array(
            '/__PHX=__GENERATE[A-B,D-O,Q-S,U-Z][A-Z]+__/',
            '/__PHX=__GENERATES[A-D,F-Z][A-Z]+__/',
            '/__PHX=__GENERATETX__/'
                ), '', $this->xmlChart);
    }

    /**
     * Generate c:txPr
     *
     * @access protected
     * @param string $font
     */
    protected function generateRFONTS2($font)
    {
        $this->xmlChart = str_replace('__PHX=__GENERATEAX__', '<c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr><a:latin typeface="' .
                $font . '" pitchFamily="34" charset="0" /><a:cs typeface="' .
                $font . '" pitchFamily="34" charset="0" /></a:defRPr></a:pPr><a:endParaRPr lang="es-ES" /></a:p></c:txPr>__PHX=__GENERATEAX__', $this->xmlChart);
    }

}