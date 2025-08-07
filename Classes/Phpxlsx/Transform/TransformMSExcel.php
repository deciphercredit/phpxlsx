<?php
namespace Phpxlsx\Transform;

use Phpxlsx\Logger\PhpxlsxLogger;

/**
 * Transform documents using MS Excel
 *
 * @category   Phpxlsx
 * @package    transform
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */

require_once dirname(__FILE__) . '/TransformPlugin.php';

class TransformMSExcel extends TransformPlugin
{
    /**
     * Transform documents:
     *     XLSX to PDF
     *
     * @access public
     * @param $source
     * @param $target
     * @param array $options
     * @return void
     */
    public function transform($source, $target, $options = array())
    {
        $allowedExtensionsSource = array('xlsx');
        $allowedExtensionsTarget = array('pdf');

        $filesExtensions = $this->checkSupportedExtension($source, $target, $allowedExtensionsSource, $allowedExtensionsTarget);

        // start an Excel instance
        $MSExcelInstance = new \COM('Excel.application') or PhpxlsxLogger::logger('Check that PHP COM is enabled and a working copy of Excel is installed.', 'fatal');

        // check that the version of MS Excel is 12 or higher
        if ($MSExcelInstance->Version >= 12) {
            // hide MS Excel
            $MSExcelInstance->Visible = 0;

            // open the source document
            $MSExcelInstance->Workbooks->Open($source);

            // save the target document
            if ($filesExtensions['targetExtension'] == 'pdf') {
                $MSExcelInstance->Workbooks[1]->ExportAsFixedFormat(0, $target, 0);
            }
        } else {
            PhpxlsxLogger::logger('The version of Excel should be 12 (Excel 2007) or higher.', 'fatal');
        }
        $MSExcelInstance->Quit();

        $MSExcelInstance = null;
        unset($MSExcelInstance);
    }
}