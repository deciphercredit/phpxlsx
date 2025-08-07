<?php
namespace Phpxlsx\Utilities;
/**
 * This class offers some utilities to work with existing Excel (.xlsx) documents
 *
 * @category   Phpxlsx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
require_once dirname(__FILE__) . '/../Create/CreateXlsx.php';

class XlsxUtilities
{
    /**
     * Search and replace shared strings and cell values in an Excel document
     *
     * @access public
     * @param string|XlsxStructure $source path to the document
     * @param string $target path to the output document
     * @param array $data strings to be searched and replaced
     * @param string scope sharedStrings, sheet
     * @param array $options
     *        sheetName : sheet name to replace the value when using sheet as scope. All if null
     *        sheetNumber : sheet number to replace the value when using sheet as scope. All if null
     * @return void
     */
    public function searchAndReplace($source, $target, $data, $scope, $options = array())
    {
        if ($source instanceof XlsxStructure) {
            // XlsxStructure object
            $xlsxFile = $source;
        } else {
            // file
            $xlsxFile = new XlsxStructure();
            $xlsxFile->parseXlsx($source);
        }

        $contentTypesXML = $xlsxFile->getContent('[Content_Types].xml');

        $contentTypesDOM = new \DOMDocument();
        if (PHP_VERSION_ID < 80000) {
            $optionEntityLoader = libxml_disable_entity_loader(true);
        }
        $contentTypesDOM->loadXML($contentTypesXML);
        if (PHP_VERSION_ID < 80000) {
            libxml_disable_entity_loader($optionEntityLoader);
        }

        $contentTypesXPath = new \DOMXPath($contentTypesDOM);
        $contentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // get current namespaces
        $namespaces = $xlsxFile->getNamespaces();

        // get application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml file
        $query = '//xmlns:Override[@ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"]';
        $mainXMLPathNodes = $contentTypesXPath->query($query);

        if ($mainXMLPathNodes->length > 0) {
            // sharedStrings contents
            if ($scope == 'sharedStrings') {
                // get application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml files
                $query = '//xmlns:Override[@ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"]';
                $sharedStringsXMLPathNodes = $contentTypesXPath->query($query);
                $sharedStringsXML = $xlsxFile->getContent(substr($sharedStringsXMLPathNodes->item(0)->getAttribute('PartName'), 1));

                $sharedStringsDOM = new \DOMDocument();
                if (PHP_VERSION_ID < 80000) {
                    $optionEntityLoader = libxml_disable_entity_loader(true);
                }
                $sharedStringsDOM->loadXML($sharedStringsXML);
                if (PHP_VERSION_ID < 80000) {
                    libxml_disable_entity_loader($optionEntityLoader);
                }

                $sharedStringsXPath = new \DOMXPath($sharedStringsDOM);
                $sharedStringsXPath->registerNamespace('xmlns', $namespaces['xmlns']);

                // replace the data
                foreach ($data as $key => $value) {
                    $this->searchToReplace($sharedStringsXPath, $key, $value);
                }

                $xlsxFile->addContent(substr($sharedStringsXMLPathNodes->item(0)->getAttribute('PartName'), 1), $sharedStringsDOM->saveXML());
            }

            // worksheet contents
            if ($scope == 'sheet') {
                // get sheets from $mainXMLPathNodes to get the correct order of the sheets
                $mainXML = $xlsxFile->getContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1));

                $mainDOM = new \DOMDocument();
                if (PHP_VERSION_ID < 80000) {
                    $optionEntityLoader = libxml_disable_entity_loader(true);
                }
                $mainDOM->loadXML($mainXML);
                if (PHP_VERSION_ID < 80000) {
                    libxml_disable_entity_loader($optionEntityLoader);
                }

                $mainXPath = new \DOMXPath($mainDOM);
                $mainXPath->registerNamespace('xmlns', $namespaces['xmlns']);

                $query = '//xmlns:sheets/xmlns:sheet';
                // query by sheet name if set
                if (isset($options['sheetName'])) {
                    $query .= '[@name="'.$options['sheetName'].'"]';
                }
                // query by sheet number if set
                if (isset($options['sheetNumber'])) {
                    $query .= '['.$options['sheetNumber'].']';
                }
                $sheetNodes = $mainXPath->query($query);

                // get sheet rels to get the sheet contents
                $mainRelsXML = $xlsxFile->getContent(str_replace('xl/', 'xl/_rels/', substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1)) . '.rels');

                $mainRelsDOM = new \DOMDocument();
                if (PHP_VERSION_ID < 80000) {
                    $optionEntityLoader = libxml_disable_entity_loader(true);
                }
                $mainRelsDOM->loadXML($mainRelsXML);
                if (PHP_VERSION_ID < 80000) {
                    libxml_disable_entity_loader($optionEntityLoader);
                }

                $mainRelsXPath = new \DOMXPath($mainRelsDOM);
                $mainRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

                $worksheetsData = array();
                foreach ($sheetNodes as $sheetNode) {
                    $query = '//xmlns:Relationship[@Id="'.$sheetNode->getAttribute('r:id').'"]';
                    $sheetContentNodes = $mainRelsXPath->query($query);
                    $worksheetsData['xl/' . $sheetContentNodes->item(0)->getAttribute('Target')] = $xlsxFile->getContent('xl/' . $sheetContentNodes->item(0)->getAttribute('Target'));
                }

                // replace the data
                foreach ($worksheetsData as $worksheetKey => $worksheetValue) {
                    $worksheetDataDOM = new \DOMDocument();
                    if (PHP_VERSION_ID < 80000) {
                        $optionEntityLoader = libxml_disable_entity_loader(true);
                    }
                    $worksheetDataDOM->loadXML($worksheetValue);
                    if (PHP_VERSION_ID < 80000) {
                        libxml_disable_entity_loader($optionEntityLoader);
                    }

                    $worksheetXPath = new \DOMXPath($worksheetDataDOM);
                    $worksheetXPath->registerNamespace('xmlns', $namespaces['xmlns']);

                    foreach ($data as $dataValue) {
                        $query = '//xmlns:sheetData/xmlns:row[@r="'.$dataValue['row'].'"]/xmlns:c['.$dataValue['col'].']/xmlns:v';
                        $dataNode = $worksheetXPath->query($query);
                        if ($dataNode->length > 0) {
                            $dataNode->item(0)->nodeValue = $dataValue['value'];
                        }
                    }

                    $worksheetsData[$worksheetKey] = $worksheetDataDOM->saveXML();
                }

                // save the data in the XLSX file
                foreach ($worksheetsData as $worksheetKey => $worksheetValue) {
                    $xlsxFile->addContent($worksheetKey, $worksheetValue);
                }
            }
        }

        // save file
        $xlsxFile->saveXlsx($target);
    }

    /**
     * Splits an Excel document
     *
     * @access public
     * @param string|XlsxStructure $source Path to the document
     * @param string $target Path to the resulting XLSX (a new file will be created per sheet)
     * @param array $options
     * @return void
     */
    public function split($source, $target, $options = array())
    {
        if ($source instanceof XlsxStructure) {
            // XlsxStructure object
            $xlsxFile = $source;
        } else {
            // file
            $xlsxFile = new XlsxStructure();
            $xlsxFile->parseXlsx($source);
        }

        $targetInfo = pathinfo($target);

        $contentTypesXML = $xlsxFile->getContent('[Content_Types].xml');

        $contentTypesDOM = new \DOMDocument();
        if (PHP_VERSION_ID < 80000) {
            $optionEntityLoader = libxml_disable_entity_loader(true);
        }
        $contentTypesDOM->loadXML($contentTypesXML);
        if (PHP_VERSION_ID < 80000) {
            libxml_disable_entity_loader($optionEntityLoader);
        }

        $contentTypesXPath = new \DOMXPath($contentTypesDOM);
        $contentTypesXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // get current namespaces
        $namespaces = $xlsxFile->getNamespaces();

        // get application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml file
        $query = '//xmlns:Override[@ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"]';
        $mainXMLPathNodes = $contentTypesXPath->query($query);

        // get sheets from $mainXMLPathNodes to get the correct order of the sheets
        $mainXML = $xlsxFile->getContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1));

        $mainDOM = new \DOMDocument();
        if (PHP_VERSION_ID < 80000) {
            $optionEntityLoader = libxml_disable_entity_loader(true);
        }
        $mainDOM->loadXML($mainXML);
        if (PHP_VERSION_ID < 80000) {
            libxml_disable_entity_loader($optionEntityLoader);
        }

        $mainXPath = new \DOMXPath($mainDOM);
        $mainXPath->registerNamespace('xmlns', $namespaces['xmlns']);

        $query = '//xmlns:sheets/xmlns:sheet';
        $sheetNodes = $mainXPath->query($query);

        // counter used for each new file name
        $i = 0;
        foreach ($sheetNodes as $sheetNode) {
            // increment the file counter
            $i++;

            // clone the source file to no overwrite it in each iteration
            $xlsxNewFile = clone $xlsxFile;

            // remove other sheets from the XLSX content
            $mainXMLNew = $xlsxFile->getContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1));
            $mainDOMNew = new \DOMDocument();
            if (PHP_VERSION_ID < 80000) {
                $optionEntityLoader = libxml_disable_entity_loader(true);
            }
            $mainDOMNew->loadXML($mainXMLNew);
            if (PHP_VERSION_ID < 80000) {
                libxml_disable_entity_loader($optionEntityLoader);
            }
            $mainXPathNew = new \DOMXPath($mainDOMNew);
            $mainXPathNew->registerNamespace('xmlns', $namespaces['xmlns']);
            $queryNew = '//xmlns:sheets/xmlns:sheet';
            $sheetNodesNew = $mainXPathNew->query($queryNew);

            // remove activeTab attribute
            $nodesWorkbookView = $mainDOMNew->getElementsByTagName('workbookView');
            if ($nodesWorkbookView->length > 0 && $nodesWorkbookView->item(0)->hasAttribute('activeTab')) {
                $nodesWorkbookView->item(0)->removeAttribute('activeTab');
            }

            $j = 1;
            foreach ($sheetNodesNew as $sheetNodeNew) {
                if ($i != $j) {
                    $sheetNodeNew->parentNode->removeChild($sheetNodeNew);
                }
                $j++;
            }

            $xlsxNewFile->addContent(substr($mainXMLPathNodes->item(0)->getAttribute('PartName'), 1), $mainDOMNew->saveXml());
            $fileXlsxPath = $targetInfo['filename'] . $i . '.' . $targetInfo['extension'];
            $xlsxNewFile->saveXlsx($fileXlsxPath);
        }
    }

    /**
     * This is the method that selects the nodes that need to be manipulated
     *
     * @access private
     * @param XPath $XPath the node to be changed
     * @return void
     */
    private function searchToReplace($xPath, $searchTerm, $replaceTerm)
    {
        $query = '//xmlns:t';
        $tNodes = $xPath->query($query);
        $searchTerm = htmlspecialchars($searchTerm);
        $replaceTerm = htmlspecialchars($replaceTerm);

        foreach ($tNodes as $tNode) {
            if (strstr($tNode->nodeValue, $searchTerm)) {
                $tNode->nodeValue = str_replace($searchTerm, $replaceTerm, $tNode->nodeValue);
            }
        }
    }
}