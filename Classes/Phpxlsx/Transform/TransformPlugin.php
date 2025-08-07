<?php
namespace Phpxlsx\Transform;

use Phpxlsx\Logger\PhpxlsxLogger;

/**
 * Transform XLSX to PDF, XLS, ODS
 *
 * @category   Phpxlsx
 * @package    transform
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */

require_once dirname(__FILE__) . '/../Create/CreateXlsx.php';

abstract class TransformPlugin
{
    /**
     *
     * @access protected
     * @var array
     */
    protected $phpxlsxconfig;

    /**
     * Construct
     *
     * @access public
     */
    public function __construct()
    {
        $this->phpxlsxconfig = \Phpxlsx\Utilities\PhpxlsxUtilities::parseConfig();
    }

    /**
     * Transform document formats
     *
     * @access public
     * @abstract
     * @param $source
     * @param $target
     * @param array $options
     */
    abstract public function transform($source, $target, $options = array());

    /**
     * Check if the extension if supproted
     *
     * @param string $fileExtension
     * @param array $supportedExtensions
     * @return array files extensions
     */
    protected function checkSupportedExtension($source, $target, $supportedExtensionsSource, $supportedExtensionsTarget) {
        // get the source file info
        $sourceFileInfo = pathinfo($source);
        $sourceExtension = strtolower($sourceFileInfo['extension']);

        if (!in_array($sourceExtension, $supportedExtensionsSource)) {
            PhpxlsxLogger::logger('The chosen extension \'' . $sourceExtension . '\' is not supported as source format.', 'fatal');
        }

        // get the target file info
        $targetFileInfo = explode('.', $target);
        $targetExtension = strtolower(array_pop($targetFileInfo));

        if (!in_array($targetExtension, $supportedExtensionsTarget)) {
            PhpxlsxLogger::logger('The chosen extension \'' . $targetExtension . '\' is not supported as target format.', 'fatal');
        }

        return array('sourceExtension' => $sourceExtension, 'targetExtension' => $targetExtension);
    }
}