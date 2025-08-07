<?php
namespace Phpxlsx\Utilities;
/**
 * Return information of an XLSX file
 *
 * @category   Phpxlsx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
require_once dirname(__FILE__) . '/../Create/CreateXlsx.php';

class Indexer
{
    /**
     * @var XlsxStructure
     */
    private $documentXlsx;

    /**
     * @var XmlUtilities XML Utilities classes
     */
    private $xmlUtilities;

    /**
     * @var array Stores the file internal structure
     */
    private $xlsxStructure;

    /**
     * Class constructor
     *
     * @param mixed $source File path or XlsxStructure
     */
    public function __construct($source)
    {
        if ($source instanceof XlsxStructure) {
            $this->documentXlsx = $source;
        } else {
            $this->documentXlsx = new XlsxStructure();
            $this->documentXlsx->parseXlsx($source);
        }

        // XmlUtilities class
        $this->xmlUtilities = new XmlUtilities();

        // init the document structure array as empty
        $this->xlsxStructure = array(
            'comments' => array(),
            'properties' => array(
                'core' => array(),
                'custom' => array(),
            ),
            'sheets' => array(),
            'signatures' => array(),
            'styles' => array(
                'cell' => array(
                    'names' => array(),
                ),
            ),
            'tables' => array(),
            'workbooks' => array(
                'sheets' => array(),
                'workbookView' => array(),
            ),
        );

        // parse the document
        $this->parse($source);
    }

    /**
     * Return a file as array or JSON
     *
     * @param string $type Output type: 'array' (default), 'json'
     * @return mixed $output
     * @throws Exception If the output type format not supported
     */
    public function getOutput($output = 'array')
    {
        // if the chosen output type is not supported throw an exception
        if (!in_array($output, array('array', 'json'))) {
            throw new \Exception('The output "' . $output . '" is not supported');
        }

        // output the document after index
        return $this->output($output);
    }

    /**
     * Extract comment contents from a XML string
     *
     * @param string $xml XML string
     */
    protected function extractComments($xml)
    {
        // check if the XML is not empty
        if (!empty($xml)) {
            // load XML content
            $contentDOM = $this->xmlUtilities->generateDomDocument($xml);
            $this->xlsxStructure['comments']['authors'] = array();
            $this->xlsxStructure['comments']['comments'] = array();

            // authors
            $authorsNodes = $contentDOM->getElementsByTagName('authors');
            if ($authorsNodes->length > 0) {
                $authorNodes = $authorsNodes->item(0)->getElementsByTagName('author');
                if ($authorNodes-> length > 0) {
                    foreach ($authorNodes as $authorNode) {
                        $this->xlsxStructure['comments']['authors'][] = $authorNode->nodeValue;
                    }
                }
            }

            // comments
            $commentListNodes = $contentDOM->getElementsByTagName('commentList');
            if ($commentListNodes->length > 0) {
                foreach ($commentListNodes as $commentListNode) {
                    $commentNodes = $contentDOM->getElementsByTagName('comment');
                    if ($commentNodes->length > 0) {
                        foreach ($commentNodes as $commentNode) {
                            $textNodes = $commentNode->getElementsByTagName('text');
                            if ($textNodes->length > 0) {
                                $textContent = '';
                                foreach ($textNodes as $textNode) {
                                    $textContent = $this->extractTexts($textNode->ownerDocument->saveXML($textNode));
                                }

                                $author = '';
                                if ($commentNode->hasAttribute('authorId')) {
                                    if (isset($this->xlsxStructure['comments']['authors']) && isset($this->xlsxStructure['comments']['authors'][$commentNode->getAttribute('authorId')])) {
                                        $author = $this->xlsxStructure['comments']['authors'][$commentNode->getAttribute('authorId')];
                                    }
                                }

                                $ref = '';
                                if ($commentNode->hasAttribute('ref')) {
                                    $ref = $commentNode->getAttribute('ref');
                                }

                                $this->xlsxStructure['comments']['comments'] = array(
                                    'author' => $author,
                                    'ref' => $ref,
                                    'text' => $textContent,
                                );
                            }
                        }
                    }
                }
            }

            // free DOMDocument resources
            $contentDOM = null;
        }
    }

    /**
     * Extract document properties from a XML string
     *
     * @param string $xml XML string
     * @param string $target Properties target: core, custom
     */
    protected function extractProperties($xml, $target)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        if ($target == 'core') {
            // do a global xpath query getting only text tags
            $contentXpath = new \DOMXPath($contentDOM);
            $contentXpath->registerNamespace('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
            $propertiesEntries = $contentXpath->query('//cp:coreProperties');

            if ($propertiesEntries->item(0)->childNodes->length > 0) {
                foreach ($propertiesEntries->item(0)->childNodes as $propertyEntry) {
                    // if empty text avoid adding the content
                    if ($propertyEntry->textContent == '') {
                        continue;
                    }

                    // get the name of the property
                    $propertyEntryFullName = explode(':', $propertyEntry->tagName);
                    $nameProperty = $propertyEntryFullName[1];

                    $this->xlsxStructure['properties']['core'][$nameProperty] = trim($propertyEntry->textContent);
                }
            }
        } else if ($target == 'custom') {
            // do a global xpath query getting only property tags
            $contentXpath = new \DOMXPath($contentDOM);
            $contentXpath->registerNamespace('ns', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties');
            $propertiesEntries = $contentXpath->query('//ns:Properties//ns:property');

            if ($propertiesEntries->length > 0) {
                foreach ($propertiesEntries as $propertyEntry) {
                    // if empty text avoid adding the content
                    if ($propertyEntry->textContent == '') {
                        continue;
                    }

                    // get the name of the property
                    $nameProperty = $propertyEntry->getAttribute('name');

                    $this->xlsxStructure['properties']['custom'][$nameProperty] = trim($propertyEntry->textContent);
                }
            }
        }

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract sheet contents from a XML string
     *
     * @param string $xml XML string
     * @param string $contentTarget Target
     */
    protected function extractSheets($xml, $contentTarget)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        $contentXPath = new \DOMXPath($contentDOM);

        // margins
        $sheetMargins = array();
        $nodesPageMargins = $contentDOM->getElementsByTagName('pageMargins');
        if ($nodesPageMargins->length > 0) {
            if ($nodesPageMargins->item(0)->hasAttributes()) {
                foreach ($nodesPageMargins->item(0)->attributes as $attribute) {
                    $sheetMargins[$attribute->nodeName] = $attribute->nodeValue;
                }
            }
        }

        // sizes
        $sheetSizes = array();
        $nodesPageSetup = $contentDOM->getElementsByTagName('pageSetup');
        if ($nodesPageSetup->length > 0) {
            if ($nodesPageSetup->item(0)->hasAttributes()) {
                foreach ($nodesPageSetup->item(0)->attributes as $attribute) {
                    $sheetSizes[$attribute->nodeName] = $attribute->nodeValue;
                }
            }
        }

        // images
        $sheetImages = array();
        $nodesDrawing = $contentDOM->getElementsByTagName('drawing');
        if ($nodesDrawing->length > 0) {
            foreach ($nodesDrawing as $nodeDrawing) {
                // rels to get the path of the drawing
                $relsFilePath = str_replace('worksheets/', 'worksheets/_rels/', $contentTarget) . '.rels';
                $contentSheetRels = $this->documentXlsx->getContent($relsFilePath);
                if (!empty($contentSheetRels)) {
                    $contentSheetRelsDom = $this->xmlUtilities->generateDomDocument($contentSheetRels);

                    $contentSheetRelsXPath = new \DOMXPath($contentSheetRelsDom);
                    $contentSheetRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                    $nodesRelationship = $contentSheetRelsXPath->query('//xmlns:Relationship[@Id="' . $nodeDrawing->getAttribute('r:id') . '"]');
                    if ($nodesRelationship->length > 0) {
                        foreach ($nodesRelationship as $nodeRelationship) {
                            $drawingTarget =  'xl/' . str_replace('../', '', $nodeRelationship->getAttribute('Target'));
                            $drawingTargetRels = str_replace('drawings/', 'drawings/_rels/', $drawingTarget) . '.rels';
                            $drawingContent = $this->documentXlsx->getContent($drawingTarget);
                            if ($drawingContent) {
                                $drawingDom = $this->xmlUtilities->generateDomDocument($drawingContent);
                                $drawingXPath = new \DOMXPath($drawingDom);
                                $drawingXPath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                                $drawingXPath->registerNamespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
                                $drawingXPath->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                                $contentDrawingRels = $this->documentXlsx->getContent($drawingTargetRels);
                                if (!empty($contentDrawingRels)) {
                                    $contentDrawingRelsDom = $this->xmlUtilities->generateDomDocument($contentDrawingRels);

                                    $contentDrawingRelsXpath = new \DOMXPath($contentDrawingRelsDom);
                                    $contentDrawingRelsXpath->registerNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships');
                                    $imageEntriesRels = $contentDrawingRelsXpath->query('//rel:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"]');

                                    if ($imageEntriesRels->length > 0) {
                                        foreach ($imageEntriesRels as $imageEntryRels) {
                                            $imageString = $this->documentXlsx->getContent('xl/' . str_replace('../', '', $imageEntryRels->getAttribute('Target')));
                                            $drawingEntries = $drawingXPath->query('//a:blip[@r:embed="'.$imageEntryRels->getAttribute('Id').'"]');

                                            if ($drawingEntries->length > 0) {
                                                foreach ($drawingEntries as $drawingEntry) {
                                                    $picNode = $drawingEntry->parentNode->parentNode;
                                                    if ($picNode) {
                                                        // init values
                                                        $heightImage = '';
                                                        $widthImage = '';
                                                        $altTextDescrImage = '';
                                                        $colPosition = array();
                                                        $rowPosition = array();

                                                        // size
                                                        $spPrNodes = $picNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'spPr');
                                                        if ($spPrNodes->length > 0)  {
                                                            foreach ($spPrNodes as $spPrNode) {
                                                                $extNodes = $spPrNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'ext');
                                                                if ($extNodes->length > 0) {
                                                                    foreach ($extNodes as $extNode) {
                                                                        if ($extNode->hasAttribute('cx')) {
                                                                            $widthImage = $extNode->getAttribute('cx');
                                                                        }
                                                                        if ($extNode->hasAttribute('cy')) {
                                                                            $heightImage = $extNode->getAttribute('cy');
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        // descr
                                                        $nvPicPrNodes = $picNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'nvPicPr');
                                                        if ($nvPicPrNodes->length > 0)  {
                                                            foreach ($nvPicPrNodes as $nvPicPrNode) {
                                                                $cNvPrNodes = $nvPicPrNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'cNvPr');
                                                                if ($cNvPrNodes->length > 0) {
                                                                    foreach ($cNvPrNodes as $cNvPrNode) {
                                                                        if ($cNvPrNode->hasAttribute('descr')) {
                                                                            $altTextDescrImage = $cNvPrNode->getAttribute('descr');
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        $twoCellAnchorNode = $drawingEntry->parentNode->parentNode->parentNode;
                                                        $positionFrom = $twoCellAnchorNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'from');
                                                        if ($positionFrom->length > 0) {
                                                            $colPosition['from'] = array();
                                                            $positionFromCol = $positionFrom->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'col');
                                                            if ($positionFromCol->length > 0) {
                                                                $colPosition['from'] = $positionFromCol->item(0)->nodeValue;
                                                            }
                                                            $rowPosition['from'] = array();
                                                            $positionFromRow = $positionFrom->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'row');
                                                            if ($positionFromRow->length > 0) {
                                                                $rowPosition['from'] = $positionFromRow->item(0)->nodeValue;
                                                            }
                                                        }
                                                        $positionTo = $twoCellAnchorNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'to');
                                                        if ($positionTo->length > 0) {
                                                            $colPosition['to'] = array();
                                                            $positionToCol = $positionTo->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'col');
                                                            if ($positionToCol->length > 0) {
                                                                $colPosition['to'] = $positionToCol->item(0)->nodeValue;
                                                            }
                                                            $rowPosition['to'] = array();
                                                            $positionToRow = $positionTo->item(0)->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'row');
                                                            if ($positionToRow->length > 0) {
                                                                $rowPosition['to'] = $positionToRow->item(0)->nodeValue;
                                                            }
                                                        }

                                                        $sheetImages[] = array(
                                                            //'content' => $imageString,
                                                            'path' => $imageEntryRels->getAttribute('Target'),
                                                            'height_word_emus' => $heightImage,
                                                            'width_word_emus' => $widthImage,
                                                            'height_word_inches' => $heightImage/914400,
                                                            'width_word_inches' => $widthImage/914400,
                                                            'height_word_cms' => $heightImage/360000,
                                                            'width_word_cms' => $widthImage/360000,
                                                            'alt_text_descr' => $altTextDescrImage,
                                                            'colPosition' => $colPosition,
                                                            'rowPosition' => $rowPosition,
                                                        );
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    // free DOMDocument resources
                                    $contentDrawingRelsDom = null;
                                }

                                // free DOMDocument resources
                                $drawingDom = null;
                            }
                        }
                    }

                    // free DOMDocument resources
                    $contentSheetRelsDom = null;
                }
            }
        }

        // links
        $sheetLinks = array();
        $nodesHyperlinks = $contentDOM->getElementsByTagName('hyperlinks');
        if ($nodesHyperlinks->length > 0) {
            $nodesHyperlink = $nodesHyperlinks->item(0)->getElementsByTagName('hyperlink');

            // rels to get the path of the hyperlink if it's an external one
            $relsFilePath = str_replace('worksheets/', 'worksheets/_rels/', $contentTarget) . '.rels';
            $contentSheetRels = $this->documentXlsx->getContent($relsFilePath);
            if (!empty($contentSheetRels)) {
                $contentSheetRelsDom = $this->xmlUtilities->generateDomDocument($contentSheetRels);
                $contentSheetRelsXPath = new \DOMXPath($contentSheetRelsDom);
                $contentSheetRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
            } else {
                $contentSheetRelsDom = null;
            }
            foreach ($nodesHyperlink as $nodeHyperlink) {
                if ($nodeHyperlink->hasAttribute('r:id')) {
                    // external link
                    if (!empty($contentSheetRelsDom)) {
                        $nodesRelationship = $contentSheetRelsXPath->query('//xmlns:Relationship[@Id="' . $nodeHyperlink->getAttribute('r:id') . '"]');
                        if ($nodesRelationship->length > 0) {
                            $target = '';
                            $ref = '';
                            if ($nodesRelationship->item(0)->hasAttribute('Target')) {
                                $target = $nodesRelationship->item(0)->getAttribute('Target');
                            }
                            if ($nodeHyperlink->hasAttribute('ref')) {
                                $ref = $nodeHyperlink->getAttribute('ref');
                            }

                            $sheetLinks[] = array(
                                'ref' => $ref,
                                'target' => $target,
                            );
                        }

                        // free DOMDocument resources
                        $contentSheetRelsDom = null;
                    }
                } else {
                    // internal link
                    $display = '';
                    $location = '';
                    $ref = '';
                    if ($nodeHyperlink->hasAttribute('display')) {
                        $display = $nodeHyperlink->getAttribute('display');
                    }
                    if ($nodeHyperlink->hasAttribute('location')) {
                        $location = $nodeHyperlink->getAttribute('location');
                    }
                    if ($nodeHyperlink->hasAttribute('ref')) {
                        $ref = $nodeHyperlink->getAttribute('ref');
                    }

                    $sheetLinks[] = array(
                        'display' => $display,
                        'location' => $location,
                        'ref' => $ref,
                    );
                }
            }
        }

        // headers and footers
        $sheetHeaders = array();
        $sheetFooters = array();
        $nodesHeaderFooter = $contentDOM->getElementsByTagName('headerFooter');
        if ($nodesHeaderFooter->length > 0) {
            foreach ($nodesHeaderFooter->item(0)->childNodes as $nodeHeaderFooter) {
                if (strstr($nodeHeaderFooter->tagName, 'Header')) {
                    $sheetHeaders[$nodeHeaderFooter->tagName] = $nodeHeaderFooter->nodeValue;
                } else if (strstr($nodeHeaderFooter->tagName, 'Footer')) {
                    $sheetFooters[$nodeHeaderFooter->tagName] = $nodeHeaderFooter->nodeValue;
                }
            }
        }

        // cells
        $sheetCells = array();
        $nodesSheetData = $contentDOM->getElementsByTagName('sheetData');
        if ($nodesSheetData->length > 0) {
            $nodesRow = $nodesSheetData->item(0)->getElementsByTagName('row');
            if ($nodesRow->length > 0) {
                foreach ($nodesRow as $nodeRow) {
                    $cellRow = array();
                    $nodesC = $nodeRow->getElementsByTagName('c');
                    if ($nodesC->length > 0) {
                        $cellC = array();
                        foreach ($nodesC as $nodeC) {
                            if ($nodeC->hasAttribute('r')) {
                                $cellC[] = $nodeC->getAttribute('r');
                            }
                        }
                    }
                    $rowValue = '';
                    if ($nodeRow->hasAttribute('r')) {
                        $rowValue = $nodeRow->hasAttribute('r');
                    }
                    $sheetCells[] = array(
                        'row' => $rowValue,
                        'cells' => $cellC,
                    );
                }
            }
        }

        $this->xlsxStructure['sheets'][] = array(
            'cells' => $sheetCells,
            'footers' => $sheetFooters,
            'headers' => $sheetHeaders,
            'images' => $sheetImages,
            'links' => $sheetLinks,
            'margins' => $sheetMargins,
            'sizes' => $sheetSizes,
        );

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract signature contents from a XML string
     *
     * @param string $xml XML string
     */
    protected function extractSignature($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        // get X509Certificate
        $contentXpath = new \DOMXPath($contentDOM);
        $contentXpath->registerNamespace('xmlns', 'http://www.w3.org/2000/09/xmldsig#');
        $x509CertificateEntry = $contentXpath->query('//xmlns:X509Certificate');
        $x509CertificateContent = null;
        if ($x509CertificateEntry->length > 0) {
            $x509Reader = openssl_x509_read("-----BEGIN CERTIFICATE-----\n" . $x509CertificateEntry->item(0)->textContent . "\n-----END CERTIFICATE-----\n");
            if ($x509Reader) {
                $x509CertificateContent = openssl_x509_parse($x509Reader);
            }
        }

        // get SignatureProperties time
        $contentXpath = new \DOMXPath($contentDOM);
        $contentXpath->registerNamespace('xmlns', 'http://www.w3.org/2000/09/xmldsig#');
        $contentXpath->registerNamespace('mdssi', 'http://schemas.openxmlformats.org/package/2006/digital-signature');
        $signatureTime = $contentXpath->query('//xmlns:SignatureProperties//mdssi:SignatureTime/mdssi:Value');
        $signatureTimeContent = null;
        if ($signatureTime->length > 0) {
            $signatureTimeContent = $signatureTime->item(0)->textContent;
        }

        // get SignatureProperties comment
        $contentXpath = new \DOMXPath($contentDOM);
        $contentXpath->registerNamespace('xmlns', 'http://www.w3.org/2000/09/xmldsig#');
        $contentXpath->registerNamespace('xmlnsdigsig', 'http://schemas.microsoft.com/office/2006/digsig');
        $signatureComment = $contentXpath->query('//xmlns:SignatureProperties//xmlnsdigsig:SignatureInfoV1//xmlnsdigsig:SignatureComments');
        $signatureCommentContent = null;
        if ($signatureComment->length > 0) {
            $signatureCommentContent = $signatureComment->item(0)->textContent;
        }

        $this->xlsxStructure['signatures'][] = array(
            'SignatureComment' => $signatureCommentContent,
            'SignatureTime' => $signatureTimeContent,
            'X509Certificate' => $x509CertificateContent,
        );

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract style contents from a XML string
     *
     * @param string $xml XML string
     */
    protected function extractStyles($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        // do a global xpath query getting only section tags
        $contentXpath = new \DOMXPath($contentDOM);
        $contentXpath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // cellStyles
        $cellStyleEntries = $contentXpath->query('//xmlns:cellStyles/xmlns:cellStyle');
        if ($cellStyleEntries->length > 0) {
            foreach ($cellStyleEntries as $cellStyleEntry) {
                if ($cellStyleEntry->hasAttribute('name')) {
                    $this->xlsxStructure['styles']['cell']['names'][] = $cellStyleEntry->getAttribute('name');
                }
            }
        }

        // cellFormats
        $cellFormatEntries = $contentXpath->query('//xmlns:cellXfs/xmlns:xf');
        if ($cellFormatEntries->length > 0) {
            foreach ($cellFormatEntries as $cellFormatEntry) {
                if ($cellFormatEntry->hasAttributes()) {
                    $attributeContents = array();
                    foreach ($cellFormatEntry->attributes as $attribute) {
                        $attributeContents[$attribute->nodeName] = $attribute->nodeValue;
                    }
                    $this->xlsxStructure['styles']['cell']['formats'][] = $attributeContents;
                }
            }
        }

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract tables contents from a XML string
     *
     * @param string $xml XML string
     */
    protected function extractTables($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        $tableInfo = array();

        // table tags
        $nodesTable = $contentDOM->getElementsByTagName('table');
        if ($nodesTable->length > 0) {
            if ($nodesTable->item(0)->hasAttribute('id')) {
                $tableInfo['id'] = $nodesTable->item(0)->getAttribute('id');
            }
            if ($nodesTable->item(0)->hasAttribute('name')) {
                $tableInfo['name'] = $nodesTable->item(0)->getAttribute('name');
            }
            if ($nodesTable->item(0)->hasAttribute('ref')) {
                $tableInfo['ref'] = $nodesTable->item(0)->getAttribute('ref');
            }
            $nodesTableColumns = $nodesTable->item(0)->getElementsByTagName('tableColumns');
            if ($nodesTableColumns->length > 0) {
                $tableInfo['columns'] = array();
                $nodesTableColumn = $nodesTableColumns->item(0)->getElementsByTagName('tableColumn');
                if ($nodesTableColumn->length > 0) {
                    foreach ($nodesTableColumn as $nodeTableColumn) {
                        $columnInfo = array();

                        if ($nodeTableColumn->hasAttribute('name')) {
                            $columnInfo['name'] = $nodeTableColumn->getAttribute('name');
                        }

                        if ($nodeTableColumn->hasAttribute('totalsRowFunction')) {
                            $columnInfo['function'] = $nodeTableColumn->getAttribute('totalsRowFunction');
                        }

                        if ($nodeTableColumn->hasAttribute('totalsRowLabel')) {
                            $columnInfo['label'] = $nodeTableColumn->getAttribute('totalsRowLabel');
                        }

                        $tableInfo['columns'][] = $columnInfo;
                    }
                }
            }
        }

        $this->xlsxStructure['tables'][] = $tableInfo;

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Extract text contents from a XML string
     *
     * @param string $xml XML string
     * @return string Text content
     */
    protected function extractTexts($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        // do a global xpath query getting only text tags
        $contentXpath = new \DOMXPath($contentDOM);
        $textEntries = $contentXpath->query('//t');

        // iterate text content and extract text strings. Add a blank space to separate each string
        $content = '';
        foreach ($textEntries as $textEntry) {
            // if empty text avoid adding the content
            if ($textEntry->textContent == ' ') {
                continue;
            }
            $content .= ' ' . $textEntry->textContent;
        }

        // free DOMDocument resources
        $contentDOM = null;

        return trim($content);
    }

    /**
     * Extract workbook contents from a XML string
     *
     * @param string $xml XML string
     */
    protected function extractWorkbooks($xml)
    {
        // load XML content
        $contentDOM = $this->xmlUtilities->generateDomDocument($xml);

        // sheets/sheet tags
        $sheetsTags = $contentDOM->getElementsByTagName('sheets');
        if ($sheetsTags->length > 0) {
            foreach ($sheetsTags as $sheetsTag) {
                $sheetTags = $sheetsTag->getElementsByTagName('sheet');
                if ($sheetTags->length > 0) {
                    foreach ($sheetTags as $sheetTag) {
                        $this->xlsxStructure['workbooks']['sheets'][] = array(
                            'name' => $sheetTag->getAttribute('name'),
                            'id' => $sheetTag->getAttribute('sheetId'),
                        );
                    }
                }
            }
        }

        // bookViews/workbookView tags
        $bookViewsTags = $contentDOM->getElementsByTagName('bookViews');
        if ($bookViewsTags->length > 0) {
            $workbookViewTags = $bookViewsTags->item(0)->getElementsByTagName('workbookView');
            if ($workbookViewTags->length > 0) {
                if ($workbookViewTags->item(0)->hasAttribute('activeTab')) {
                    $this->xlsxStructure['workbooks']['workbookView']['activeTab'] = $workbookViewTags->item(0)->getAttribute('activeTab');
                }
            }
        }

        // definedNames/definedName tags
        $definedNamesTags = $contentDOM->getElementsByTagName('definedNames');
        if ($definedNamesTags->length > 0) {
            $definedNamesTags = $definedNamesTags->item(0)->getElementsByTagName('definedName');
            if ($definedNamesTags->length > 0) {
                $this->xlsxStructure['workbooks']['definedNames'] = array();
                foreach ($definedNamesTags as $definedNamesTag) {
                    $this->xlsxStructure['workbooks']['definedNames'][] = array(
                        'name' => $definedNamesTag->getAttribute('name'),
                        'value' => $definedNamesTag->nodeValue,
                        'comment' => $definedNamesTag->getAttribute('comment'),
                    );
                }
                if ($workbookViewTags->item(0)->hasAttribute('activeTab')) {
                    $this->xlsxStructure['workbooks']['workbookView']['activeTab'] = $workbookViewTags->item(0)->getAttribute('activeTab');
                }
            }
        }

        // free DOMDocument resources
        $contentDOM = null;
    }

    /**
     * Return a file as array or JSON
     *
     * @param string $type Output type: 'array' (default), 'json'
     * @return mixed $output
     */
    protected function output($type = 'array')
    {
        // array as default
        $output = $this->xlsxStructure;

        // export as the choosen type
        if ($type == 'json') {
            $output = json_encode($output);
        }

        return $output;
    }

    /**
     * Parse an XLSX file
     *
     * @param XLSXStructure $source
     */
    private function parse($source)
    {
        // parse the Content_Types
        $contentTypesContent = $this->documentXlsx->getContent('[Content_Types].xml');
        $contentTypesXml = simplexml_load_string($contentTypesContent);

        $contentTypesDom = dom_import_simplexml($contentTypesXml);

        $contentTypesXpath = new \DOMXPath($contentTypesDom->ownerDocument);
        $contentTypesXpath->registerNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/content-types');
        $relsEntries = $contentTypesXpath->query('//rel:Default[@ContentType="application/vnd.openxmlformats-package.relationships+xml"]');
        $relsExtension = 'rels';
        if (isset($relsEntries[0])) {
            $relsExtension = $relsEntries[0]->getAttribute('Extension');
        }

        // iterate over the Content_Types
        foreach ($contentTypesXml->Override as $override) {
            foreach ($override->attributes() as $attribute => $value) {
                // get the file content
                $contentTarget = substr($override->attributes()->PartName, 1);
                $content = $this->documentXlsx->getContent($contentTarget);

                // before adding a content remove the first character to get the right file path
                // removing the first slash of each path
                if ($value == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml') {
                    // workbook content

                    // extract workbook
                    $this->extractWorkbooks($content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml') {
                    // sheet content

                    // extract sheet
                    $this->extractSheets($content, $contentTarget);
                } else if ($value == 'application/vnd.openxmlformats-package.core-properties+xml') {
                    // core properties content

                    // extract core properties
                    $this->extractProperties($content, 'core');
                } else if ($value == 'application/vnd.openxmlformats-officedocument.custom-properties+xml') {
                    // custom properties content

                    // extract custom properties
                    $this->extractProperties($content, 'custom');
                } else if ($value == 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml') {
                    // signature contents

                    // extract signatures
                    $this->extractSignature($content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml') {
                    // comment contents

                    // extract comments
                    $this->extractComments($content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml') {
                    // table contents

                    // extract tables
                    $this->extractTables($content);
                } else if ($value == 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml') {
                    // style contents

                    // extract tables
                    $this->extractStyles($content);
                }
            }
        }
    }
}