<?php
namespace Phpxlsx\Utilities;
/**
 * XML functions
 *
 * @category   Phpxlsx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class XmlUtilities
{
    /**
     * Generate a DOM document from a XML string
     * @param string $xml XML content
     */
    public function generateDOMDocument($xml)
    {
        $domDocument = new \DOMDocument();
        if (PHP_VERSION_ID < 80000) {
            $optionEntityLoader = libxml_disable_entity_loader(true);
        }
        $domDocument->loadXML($xml);
        if (PHP_VERSION_ID < 80000) {
            libxml_disable_entity_loader($optionEntityLoader);
        }

        return $domDocument;
    }
}