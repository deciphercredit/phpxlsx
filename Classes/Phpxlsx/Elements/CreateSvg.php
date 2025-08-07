<?php
namespace Phpxlsx\Elements;
/**
 * Create SVG
 *
 * @category   Phpxlsx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class CreateSvg extends CreateElement
{
    /**
     * Scale image factor
     */
    const SCALEFACTOR = 360000;

    /**
     * Create svg
     *
     * @access public
     * @param string $svg
     * @param string $image
     * @param array $position
     * @param array $options
     *      'colOffset' (array) given in emus (1cm = 360000 emus). 0 as default
     *          'from' (int) from offset
     *          'to' (int) from offset
     *      'colSize' (int) number of cols used by the image
     *      'descr' (string) set a descr value
     *      'dpi' (int) dots per inch
     *      'editAs' (string) oneCell (default) (move but don't size with cells), twoCell (move and size with cells), absolute (don't move or size with cells)
     *      'imageInformation' (array) width and height values
     *      'name' (string) set a name value
     *      'rIdSVG' (string) SVG ID
     *      'rIdAlt' (string) alt image ID
     *      'rowOffset' (array) given in emus (1cm = 360000 emus). 0 as default
     *          'from' (int) from offset
     *          'to' (int) from offset
     *      'rowSize' (int) number of rows used by the image
     * @return array
     */
    public function createElementSvg($svg, $image, $position, $options = array())
    {
        // from and to column values
        $colFromValue = $this->wordToInt($position['text']);
        $colToValue = $colFromValue + 1;
        if (isset($options['colSize'])) {
            $colToValue = $colFromValue + $options['colSize'];
        }
        $colFromOffsetValue = 0;
        if (isset($options['colOffset']) && isset($options['colOffset']['from'])) {
            $colFromOffsetValue = $options['colOffset']['from'];
        }
        $colToOffsetValue = 0;
        if (isset($options['colOffset']) && isset($options['colOffset']['to'])) {
            $colToOffsetValue = $options['colOffset']['to'];
        }

        // from and to row values
        $rowFromValue = $position['number'];
        // start from 0
        $rowFromValue--;
        $rowToValue = $rowFromValue + 1;
        if (isset($options['rowSize'])) {
            $rowToValue = $rowFromValue + $options['rowSize'];
        }
        $rowFromOffsetValue = 0;
        if (isset($options['rowOffset']) && isset($options['rowOffset']['from'])) {
            $rowFromOffsetValue = $options['rowOffset']['from'];
        }
        $rowToOffsetValue = 0;
        if (isset($options['rowOffset']) && isset($options['rowOffset']['to'])) {
            $rowToOffsetValue = $options['rowOffset']['to'];
        }

        $newContents = array(
            'drawingXml' => '',
        );

        $width = null;
        if (isset($options['imageInformation']['width'])) {
            $width = $options['imageInformation']['width'];
            if (isset($options['scaling'])) {
                $width = $width * $options['scaling'] / 100;
            }
        }

        $height = null;
        if (isset($options['imageInformation']['height'])) {
            $height = $options['imageInformation']['height'];
            if (isset($options['scaling'])) {
                $height = $height * $options['scaling'] / 100;
            }
        }

        $dpi = 96;
        if (isset($options['dpi'])) {
            $dpi = $options['dpi'];
        }

        $name = 'Picture ' . $options['rIdSVG'];
        if (isset($options['name'])) {
            $name = $this->parseAndCleanTextString($options['name']);
        }

        $descr = '';
        if (isset($options['descr'])) {
            $descr = $this->parseAndCleanTextString($options['descr']);
        }

        $width = round($width * 2.54 / $dpi * self::SCALEFACTOR);
        $height = round($height * 2.54 / $dpi * self::SCALEFACTOR);

        $cNvPrId = rand(999999, 999999999);

        $newContents['drawingXml'] = '<xdr:twoCellAnchor editAs="'.$options['editAs'].'" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:from><xdr:col>'.$colFromValue.'</xdr:col><xdr:colOff>'.$colFromOffsetValue.'</xdr:colOff><xdr:row>'.$rowFromValue.'</xdr:row><xdr:rowOff>'.$rowFromOffsetValue.'</xdr:rowOff></xdr:from><xdr:to><xdr:col>'.$colToValue.'</xdr:col><xdr:colOff>'.$colToOffsetValue.'</xdr:colOff><xdr:row>'.$rowToValue.'</xdr:row><xdr:rowOff>'.$rowToOffsetValue.'</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr descr="'.$descr.'" id="'.$cNvPrId.'" name="'.$name.'"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId'.$options['rIdAlt'].'"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi val="0" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"/></a:ext><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId'.$options['rIdSVG'].'" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"/></a:ext></a:extLst></a:blip><a:stretch/></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'.$width.'" cy="'.$height.'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="0"><a:noFill/></a:ln></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>';

        return $newContents;
    }
}