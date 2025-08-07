<?php
namespace Phpxlsx\Transform;

use Phpxlsx\Logger\PhpxlsxLogger;

/**
 * Transform documents using LibreOffice
 *
 * @category   Phpxlsx
 * @package    transform
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */

require_once dirname(__FILE__) . '/TransformPlugin.php';

class TransformLibreOffice extends TransformPlugin
{
    /**
     * Transform:
     *     XLSX to PDF, XLS, ODS
     *     XLS to XLSX, PDF, ODS
     *     ODS to XLSX, PDF, XLS
     *
     * @access public
     * @param $source
     * @param $target
     * @param array $options :
     *   'debug' (bool) : false (default) or true. Shows debug information about the plugin conversion
     *   'extraOptions' (string) : extra parameters to be used when doing the conversion
     *   'homeFolder' (string) : set a custom home folder to be used for the conversions
     *   'outdir' (string) : set the outdir path. Useful when the PDF output path is not the same than the running script
     * @return void
     */
    public function transform($source, $target, $options = array())
    {
        $allowedExtensionsSource = array('xls', 'xlsx', 'ods');
        $allowedExtensionsTarget = array('xls', 'xlsx', 'pdf', 'ods');

        $filesExtensions = $this->checkSupportedExtension($source, $target, $allowedExtensionsSource, $allowedExtensionsTarget);

        if (!isset($options['debug'])) {
            $options['debug'] = false;
        }
        // rename the output file to target as default
        $renameFiles = true;

        // get the file info
        $sourceFileInfo = pathinfo($source);
        $sourceExtension = $sourceFileInfo['extension'];

        $phpxlsxconfig = \Phpxlsx\Utilities\PhpxlsxUtilities::parseConfig();
        $libreOfficePath = $phpxlsxconfig['transform']['path'];

        $customHomeFolder = false;
        if (isset($options['homeFolder'])) {
            $currentHomeFolder = getenv("HOME");
            putenv("HOME=" . $options['homeFolder']);
            $customHomeFolder = true;
        } else if (isset($phpxlsxconfig['transform']['home_folder'])) {
            $currentHomeFolder = getenv("HOME");
            putenv("HOME=" . $phpxlsxconfig['transform']['home_folder']);
            $customHomeFolder = true;
        }

        $extraOptions = '';
        if (isset($options['extraOptions'])) {
            $extraOptions = $options['extraOptions'];
        }

        // set outputstring for debugging
        $outputDebug = '';
        if (PHP_OS == 'Linux' || PHP_OS == 'Darwin' || PHP_OS == ' FreeBSD') {
            if (!$options['debug']) {
                $outputDebug = ' > /dev/null 2>&1';
            }
        } elseif (substr(PHP_OS, 0, 3) == 'Win' || substr(PHP_OS, 0, 3) == 'WIN') {
            if (!$options['debug']) {
                $outputDebug = ' > nul 2>&1';
            }
        }

        // if the outdir option is set use it as target path, instead use the dir path
        if (isset($options['outdir'])) {
            $outdir = $options['outdir'];
        } else {
            $outdir = $sourceFileInfo['dirname'];
        }

        // call LibreOffice
        passthru($libreOfficePath . ' ' . $extraOptions . ' --invisible --convert-to ' . $filesExtensions['targetExtension'] . ' ' . $source . ' --outdir ' . $outdir . $outputDebug);

        // get the converted document, this is the name of the source and the extension
        $newDocumentPath = $outdir . '/' . $sourceFileInfo['filename'] . '.' . $filesExtensions['targetExtension'];

        // move the document to the guessed destination
        if ($renameFiles) {
            rename($newDocumentPath, $target);
        }

        // restore the previous HOME value if a custom one has been set
        if ($customHomeFolder) {
            putenv("HOME=" . $currentHomeFolder);
        }
    }
}