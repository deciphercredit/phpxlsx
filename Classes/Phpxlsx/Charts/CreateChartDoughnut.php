<?php
namespace Phpxlsx\Charts;

use Phpxlsx\Elements\CreateChartElement;

/**
 * Create doughnut chart
 *
 * @category   Phpxlsx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class CreateChartDoughnut extends CreateChartElement
{
    /**
     * Create embedded xml chart
     *
     * @access public
     */
    public function createEmbeddedXmlChart()
    {
        $this->xmlChart = '';
        $this->generateCHARTSPACE();
        $this->generateDATE1904(1);
        $this->generateLANG();
        $this->generateROUNDEDCORNERS($this->roundedCorners);
        $color = 2;
        if (!empty($this->color)) {
            $color = $this->color;
        }
        $this->generateSTYLE($color);
        $this->generateCHART();
        if ($this->title != '') {
            $this->generateTITLE();
            $this->generateTITLETX();
            $this->generateRICH();
            $this->generateBODYPR();
            $this->generateLSTSTYLE();
            $this->generateTITLEP();
            $this->generateTITLEPPR();
            $this->generateDEFRPR('title');
            $this->generateTITLER();
            $this->generateTITLERPR();
            $this->generateTITLET($this->title);
            $this->cleanTemplateFonts();
        } else {
            $this->generateAUTOTITLEDELETED();
            $title = '';
        }

        if (strpos($this->type, '3D') !== false) {
            $this->generateVIEW3D();
            $rotX = 30;
            $rotY = 30;
            $perspective = 30;
            if ($this->rotX != '') {
                $rotX = $this->rotX;
            }
            if ($this->rotY != '') {
                $rotY = $this->rotY;
            }
            if ($this->perspective != '') {
                $perspective = $this->perspective;
            }
            $this->generateROTX($rotX);
            $this->generateROTY($rotY);
            $this->generateRANGAX($this->rAngAx);
            $this->generatePERSPECTIVE($perspective);
        }
        $this->generatePLOTAREA();
        $this->generateLAYOUT();

        $this->generateDOUGHNUTCHART();
        $this->generateVARYCOLORS();
        $legends = array($this->title);
        for ($i = 0; $i < count($legends); $i++) {
            $this->generateSER();
            $this->generateIDX($i);
            $this->generateORDER($i);

            $this->generateTX();
            $this->generateSTRREF();
            //$this->generateF('Sheet1!$' . $letter . '$1');
            $this->generateSTRCACHE();
            $this->generatePTCOUNT();
            $this->generatePT();
            $this->generateV($legends[$i]);
            if (!empty($this->explosion) && is_numeric($this->explosion)) {
                $this->generateEXPLOSION($this->explosion);
            }
            $this->cleanTemplate2();

            if (is_array($this->theme) && isset($this->theme['serRgbColors']) && isset($this->theme['serRgbColors'][$i])) {
                if ($this->theme['serRgbColors'][$i] != null) {
                    $this->generateSPPR_SER();
                    $this->generateSPPR_SOLIDFILL($this->theme['serRgbColors'][$i]);
                }
            }

            if (is_array($this->theme) && isset($this->theme['valueRgbColors']) && isset($this->theme['valueRgbColors'][$i]) && $this->theme['valueRgbColors'][$i] != null) {
                if ($this->theme['valueRgbColors'][$i] != null) {
                    $this->generateCDPT($this->theme['valueRgbColors'][$i]);
                }
            }

            $this->generateCAT();
            $this->generateSTRREF();
            $this->generateF($this->options['legendsRef']);
            $this->generateSTRCACHE();
            $this->generatePTCOUNT(count($this->values['contentLegends']));

            $num = 0;
            foreach ($this->values['contentLegends'] as $value) {
                $this->generatePT($num);
                $this->generateV($value);
                $num++;
            }
            $this->cleanTemplate2();
            $this->generateVAL();
            $this->generateNUMREF();
            $this->generateF($this->options['valuesRef']);
            $this->generateNUMCACHE();
            $this->generateFORMATCODE();
            $this->generatePTCOUNT(count($this->values['contentValues']));
            $num = 0;
            foreach ($this->values['contentValues'] as $value) {
                $this->generatePT($num);
                $this->generateV($value);
                $num++;
            }
            $this->cleanTemplate3();
        }

        //Generate labels
        $this->generateSERDLBLS();

        if ($this->formatCode) {
            $this->generateNUMFMT($this->formatCode, 0);
        }

        $this->generateSHOWLEGENDKEY($this->showLegendKey);
        $this->generateSHOWVAL($this->showValue);
        $this->generateSHOWCATNAME($this->showCategory);
        $this->generateSHOWSERNAME($this->showSeries);
        $this->generateSHOWPERCENT($this->showPercent);
        $this->generateSHOWBUBBLESIZE($this->showBubbleSize);
        $this->generateFIRSTSLICEANG();
        if (!empty($this->holeSize) && is_numeric($this->holeSize)) {
            $this->generateHOLESIZE($this->holeSize);
        } else {
            $this->generateHOLESIZE();
        }

        $this->generateLEGEND();
        $this->generateLEGENDPOS($this->legendPos);
        $this->generateLEGENDOVERLAY($this->legendOverlay);
        $this->generatePLOTVISONLY();

        if ((!isset($this->border) || $this->border == 0 || !is_numeric($this->border))
        ) {
            $this->generateSPPR();
            $this->generateLN();
            $this->generateNOFILL();
        } else {
            $this->generateSPPR();
            $this->generateLN($this->border);
        }

        if ($this->font != '') {
            $this->generateTXPR();
            $this->generateLEGENDBODYPR();
            $this->generateLSTSTYLE();
            $this->generateAP();
            $this->generateAPPR();
            $this->generateDEFRPR();
            $this->generateRFONTS($this->font);
            $this->generateENDPARARPR();
        }

        $this->cleanTemplateDocument();
        return $this->xmlChart;
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
        $this->xmlChart = str_replace('__PHX=__GENERATEDLBLS__', '<c:numFmt formatCode="' . $formatCode . '" sourceLinked="' . $sourceLinked . '"></c:numFmt>__PHX=__GENERATEDLBLS__', $this->xmlChart);
    }
}