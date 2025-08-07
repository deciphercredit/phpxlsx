<?php
namespace Phpxlsx\Create;

use Phpxlsx\Logger\PhpxlsxLogger;
use Phpxlsx\Resources\OOXMLResources;
use Phpxlsx\Utilities\ImageUtilities;
use Phpxlsx\Utilities\XlsxStructure;

/**
 * Use an existing XLSX as the document template
 *
 * @category   Phpxlsx
 * @package    create
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class CreateXlsxFromTemplate extends CreateXlsx
{
    /**
     *
     * @access public
     * @var string
     */
    public $templateSymbolStart = '$';

    /**
     *
     * @access public
     * @var string
     */
    public $templateSymbolEnd = '$';

    /**
     *
     * @access public
     * @static
     * @var string
     */
    public static $regExprVariableSymbols = '([^ ]*)';

    /**
     *
     * @access public
     * @static
     * @var string
     */
    public static $templateBlockSymbol = 'BLOCK_';

    /**
     * Construct
     *
     * @access public
     * @param mixed $xlsxTemplatePath path to the template to use or XlsxStructure
     * @param array $options
     * @throws Exception empty or not valid template
     */
    public function __construct($xlsxTemplatePath, $options = array())
    {
        if (empty($xlsxTemplatePath)) {
            PhpxlsxLogger::logger('The template path can not be empty.', 'fatal');
        }
        parent::__construct($options, $xlsxTemplatePath);
    }

    /**
     * Getter. Return template symbol
     *
     * @access public
     * @return string|array
     */
    public function getTemplateSymbol()
    {
        if ($this->templateSymbolStart == $this->templateSymbolEnd) {
            return $this->templateSymbolStart;
        } else {
            return array($this->templateSymbolStart, $this->templateSymbolEnd);
        }
    }

    /**
     * Setter. Set template symbol
     *
     * @access public
     * @param string $templateSymbolStart
     * @param string $templateSymbolEnd use $templateSymbolStart if null
     */
    public function setTemplateSymbol($templateSymbolStart = '$', $templateSymbolEnd = null)
    {
        if (is_null($templateSymbolEnd)) {
            $this->templateSymbolStart = $templateSymbolStart;
            $this->templateSymbolEnd = $templateSymbolStart;
        } else {
            $this->templateSymbolStart = $templateSymbolStart;
            $this->templateSymbolEnd = $templateSymbolEnd;
        }

        PhpxlsxLogger::logger('Set new template symbol.', 'info');
    }

    /**
     * Returns the template variables
     *
     * @access public
     * @param string $target may be all (default), sheets, headers, footers or comments
     * @param array $variables
     * @return array
     */
    public function getTemplateVariables($target = 'all', $variables = array())
    {
        $targetTypes = array('sheets', 'headers', 'footers', 'comments');

        if ($target == 'sheets') {
            // strings
            $sharedStringsContents = $this->zipXlsx->getContentByType('sharedStrings');
            foreach ($sharedStringsContents as $sharedStringsContent) {
                // iterate t tags
                $sharedStringsDOM = $this->xmlUtilities->generateDomDocument($sharedStringsContent['content']);

                $nodesT = $sharedStringsDOM->getElementsByTagName('t');

                foreach ($nodesT as $nodeT) {
                    $newVariables = $this->extractVariables($nodeT->nodeValue);
                    foreach ($newVariables as $newVariable) {
                        if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                            $variables[$target][] = $newVariable;
                        }
                    }
                }

                // free DOMDocument resources
                $sharedStringsDOM = null;
            }

            // drawings
            $sheetsContents = $this->zipXlsx->getSheets();
            foreach ($sheetsContents as $sheetContents) {
                $sheetDOM = $this->xmlUtilities->generateDomDocument($sheetContents['content']);
                $sheetXPath = new \DOMXPath($sheetDOM);
                $sheetXPath->registerNamespace('xmlns', $this->namespaces['xmlns']);
                $queryContents = '//xmlns:drawing|//xmlns:legacyDrawing';
                $nodesDrawing = $sheetXPath->query($queryContents);

                foreach ($nodesDrawing as $nodeDrawing) {
                    // rels to get the path of the drawing
                    $relsFilePath = str_replace('worksheets/', 'worksheets/_rels/', $sheetContents['path']) . '.rels';
                    $sheetRelsContent = $this->zipXlsx->getContent($relsFilePath);
                    if (!empty($sheetRelsContent)) {
                        $sheetRelsDOM = $this->xmlUtilities->generateDomDocument($sheetRelsContent);

                        $sheetRelsXPath = new \DOMXPath($sheetRelsDOM);
                        $sheetRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                        $nodesRelationship = $sheetRelsXPath->query('//xmlns:Relationship[@Id="' . $nodeDrawing->getAttribute('r:id') . '"]');
                        if ($nodesRelationship->length > 0) {
                            foreach ($nodesRelationship as $nodeRelationship) {
                                $drawingTarget =  'xl/' . str_replace('../', '', $nodeRelationship->getAttribute('Target'));
                                $drawingContent = $this->zipXlsx->getContent($drawingTarget);
                                if ($drawingContent) {
                                    $drawingDOM = $this->xmlUtilities->generateDomDocument($drawingContent);
                                    $nodesCNVPR = $drawingDOM->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'cNvPr');
                                    if ($nodesCNVPR->length > 0) {
                                        foreach ($nodesCNVPR as $nodeCNVPR) {
                                            $nodeCNVPRDescrValue = $nodeCNVPR->getAttribute('descr');
                                            if ($nodeCNVPR->hasAttribute('descr') && !empty($nodeCNVPRDescrValue)) {
                                                $newVariables = $this->extractVariables($nodeCNVPR->getAttribute('descr'));
                                                foreach ($newVariables as $newVariable) {
                                                    if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                                                        $variables[$target][] = $newVariable;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // free DOMDocument resources
                $sheetDOM = null;
            }
        } else if ($target == 'headers' || $target == 'footers') {
            // sheets
            $sheetsContents = $this->zipXlsx->getSheets();
            foreach ($sheetsContents as $sheetContents) {
                $sheetDOM = $this->xmlUtilities->generateDomDocument($sheetContents['content']);
                $sheetXPath = new \DOMXPath($sheetDOM);
                $sheetXPath->registerNamespace('xmlns', $this->namespaces['xmlns']);
                $queryContents = '';
                if ($target == 'headers') {
                    $queryContents = '//xmlns:headerFooter/xmlns:oddHeader|//xmlns:headerFooter/xmlns:evenHeader|//xmlns:headerFooter/xmlns:firstHeader';
                } else if ($target == 'footers') {
                    $queryContents = '//xmlns:headerFooter/xmlns:oddFooter|//xmlns:headerFooter/xmlns:evenFooter|//xmlns:headerFooter/xmlns:firstFooter';
                }
                $headersFootersContents = $sheetXPath->query($queryContents);
                if ($headersFootersContents->length > 0) {
                    foreach ($headersFootersContents as $headersFootersContent) {
                        $newVariables = $this->extractVariables($headersFootersContent->ownerDocument->saveXML($headersFootersContent));
                        foreach ($newVariables as $newVariable) {
                            if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                                $variables[$target][] = $newVariable;
                            }
                        }
                    }
                }

                // free DOMDocument resources
                $sheetDOM = null;
            }

            // drawings
            foreach ($sheetsContents as $sheetContents) {
                $sheetDOM = $this->xmlUtilities->generateDomDocument($sheetContents['content']);
                $sheetXPath = new \DOMXPath($sheetDOM);
                $sheetXPath->registerNamespace('xmlns', $this->namespaces['xmlns']);
                $queryContents = '//xmlns:drawingHF|//xmlns:legacyDrawingHF';
                $nodesDrawing = $sheetXPath->query($queryContents);

                foreach ($nodesDrawing as $nodeDrawing) {
                    // rels to get the path of the drawing
                    $relsFilePath = str_replace('worksheets/', 'worksheets/_rels/', $sheetContents['path']) . '.rels';
                    $sheetRelsContent = $this->zipXlsx->getContent($relsFilePath);
                    if (!empty($sheetRelsContent)) {
                        $sheetRelsDOM = $this->xmlUtilities->generateDomDocument($sheetRelsContent);

                        $sheetRelsXPath = new \DOMXPath($sheetRelsDOM);
                        $sheetRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                        $nodesRelationship = $sheetRelsXPath->query('//xmlns:Relationship[@Id="' . $nodeDrawing->getAttribute('r:id') . '"]');
                        if ($nodesRelationship->length > 0) {
                            foreach ($nodesRelationship as $nodeRelationship) {
                                $drawingTarget =  'xl/' . str_replace('../', '', $nodeRelationship->getAttribute('Target'));
                                $drawingContent = $this->zipXlsx->getContent($drawingTarget);
                                if ($drawingContent) {
                                    $drawingDOM = $this->xmlUtilities->generateDomDocument($drawingContent);
                                    $nodesShape = $drawingDOM->getElementsByTagNameNS('urn:schemas-microsoft-com:vml', 'shape');
                                    if ($nodesShape->length > 0) {
                                        foreach ($nodesShape as $nodeShape) {
                                            $nodeShapeAltValue = $nodeShape->getAttribute('alt');
                                            if ($nodeShape->hasAttribute('alt') && !empty($nodeShapeAltValue)) {
                                                $newVariables = $this->extractVariables($nodeShape->getAttribute('alt'));
                                                // headers scope
                                                if ($target == 'headers' && $nodeShape->hasAttribute('id') && in_array($nodeShape->getAttribute('id'), XlsxStructure::$idsHeaders)) {
                                                    foreach ($newVariables as $newVariable) {
                                                        if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                                                            $variables[$target][] = $newVariable;
                                                        }
                                                    }
                                                }
                                                // footers scope
                                                if ($target == 'footers' && $nodeShape->hasAttribute('id') && in_array($nodeShape->getAttribute('id'), XlsxStructure::$idsFooters)) {
                                                    foreach ($newVariables as $newVariable) {
                                                        if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                                                            $variables[$target][] = $newVariable;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    $nodesDrawing = $drawingDOM->getElementsByTagNameNS('http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing', 'cNvPr');
                                    if ($nodesDrawing->length > 0) {
                                        foreach ($nodesDrawing as $nodeDrawing) {
                                            $nodeDescrValue = $nodeDrawing->getAttribute('descr');
                                            if ($nodeDrawing->hasAttribute('descr') && !empty($nodeDescrValue)) {
                                                $newVariables = $this->extractVariables($nodeDrawing->getAttribute('descr'));
                                                // headers scope
                                                if ($target == 'headers' && $nodeDrawing->hasAttribute('name') && in_array($nodeDrawing->getAttribute('name'), XlsxStructure::$idsHeaders)) {
                                                    foreach ($newVariables as $newVariable) {
                                                        if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                                                            $variables[$target][] = $newVariable;
                                                        }
                                                    }
                                                }
                                                // footers scope
                                                if ($target == 'footers' && $nodeDrawing->hasAttribute('name') && in_array($nodeDrawing->getAttribute('name'), XlsxStructure::$idsFooters)) {
                                                    foreach ($newVariables as $newVariable) {
                                                        if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                                                            $variables[$target][] = $newVariable;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    // free DOMDocument resources
                                    $drawingDOM = null;
                                }
                            }
                        }

                        // free DOMDocument resources
                        $sheetRelsDOM = null;
                    }
                }
            }
        } else if ($target == 'comments') {
            // comments
            $commentsContents = $this->zipXlsx->getContentByType('comments');
            foreach ($commentsContents as $commentContents) {
                // iterate t tags
                $sheetDOM = $this->xmlUtilities->generateDomDocument($commentContents['content']);

                $nodesT = $sheetDOM->getElementsByTagName('t');

                foreach ($nodesT as $nodeT) {
                    $newVariables = $this->extractVariables($nodeT->nodeValue);
                    foreach ($newVariables as $newVariable) {
                        if (!isset($variables[$target]) || !in_array($newVariable, $variables[$target])) {
                            $variables[$target][] = $newVariable;
                        }
                    }
                }

                // free DOMDocument resources
                $sheetDOM = null;
            }
        } else if ($target == 'all') {
            foreach ($targetTypes as $targets) {
                $variables = $this->getTemplateVariables($targets, $variables);
            }
        }

        PhpxlsxLogger::logger('Get template variables.', 'info');

        return $variables;
    }

    /**
     * Removes template text variables
     *
     * @access public
     * @param array $variables Variables to be removed
     * @param array $options
     *      'target' (string) sheets (default), headers, footers, comments
     */
    public function removeVariableText($variables, $options = array())
    {
        // set empty values for all variables to be used with replaceVariableText
        $variablesFilled = array();
        foreach ($variables as $variable) {
            $variablesFilled[$variable] = '';
        }

        PhpxlsxLogger::logger('Remove template variable.', 'info');

        $this->replaceVariableText($variablesFilled, $options);
    }

    /**
     * Replaces image placeholders by an external image
     *
     * @access public
     * @param string $variables variable names and image paths
     * @param array $options
     *      'mime' (string) forces a mime (image/jpg, image/jpeg, image/png, image/gif)
     *      'target' (string) sheets (default), headers, footers
     * @throws \Exception image doesn't exist
     * @throws \Exception image format is not supported
     */
    public function replaceVariableImage($variables, $options = array())
    {
        if (!isset($options['target'])) {
            $options['target'] = 'sheets';
        }

        if ($options['target'] == 'sheets' || $options['target'] == 'headers' || $options['target'] == 'footers') {
            $sheetsContents = $this->zipXlsx->getSheets();
            foreach ($sheetsContents as $sheetContents) {
                // get drawing tags that include the images
                $sheetDOM = $this->xmlUtilities->generateDomDocument($sheetContents['content']);
                $sheetXPath = new \DOMXPath($sheetDOM);
                $sheetXPath->registerNamespace('xmlns', $this->namespaces['xmlns']);

                // drawing tags
                if ($options['target'] == 'sheets') {
                    // sheets use drawing and legacyDrawing tags
                    $queryContents = '//xmlns:drawing|//xmlns:legacyDrawing';
                    $nodesDrawing = $sheetXPath->query($queryContents);
                } else if ($options['target'] == 'headers' || $options['target'] == 'footers') {
                    // headers and footers use drawingHF and legacyDrawingHF tags
                    $queryContents = '//xmlns:drawingHF|//xmlns:legacyDrawingHF';
                    $nodesDrawing = $sheetXPath->query($queryContents);
                }
                if ($nodesDrawing->length > 0) {
                    foreach ($nodesDrawing as $nodeDrawing) {
                        // rels to get the path of the drawing
                        $relsFilePath = str_replace('worksheets/', 'worksheets/_rels/', $sheetContents['path']) . '.rels';
                        $sheetRelsContent = $this->zipXlsx->getContent($relsFilePath);
                        if (!empty($sheetRelsContent)) {
                            $sheetRelsDOM = $this->xmlUtilities->generateDomDocument($sheetRelsContent);

                            $sheetRelsXPath = new \DOMXPath($sheetRelsDOM);
                            $sheetRelsXPath->registerNamespace('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
                            $nodesRelationship = $sheetRelsXPath->query('//xmlns:Relationship[@Id="' . $nodeDrawing->getAttribute('r:id') . '"]');
                            if ($nodesRelationship->length > 0) {
                                foreach ($nodesRelationship as $nodeRelationship) {
                                    $drawingTarget =  'xl/' . str_replace('../', '', $nodeRelationship->getAttribute('Target'));
                                    $drawingTargetRels = str_replace('drawings/', 'drawings/_rels/', $drawingTarget) . '.rels';
                                    $drawingContent = $this->zipXlsx->getContent($drawingTarget);
                                    if ($drawingContent) {
                                        $drawingDOM = $this->xmlUtilities->generateDomDocument($drawingContent);
                                        $drawingRelsContent = $this->zipXlsx->getContent($drawingTargetRels);
                                        if (empty($drawingRelsContent)) {
                                            $drawingRelsContent = OOXMLResources::$sheetRelsXML;
                                        }
                                        $drawingRelsDOM = $this->xmlUtilities->generateDomDocument($drawingRelsContent);

                                        foreach ($variables as $variableKey => $variableSrc) {
                                            $newRelsDrawing = $this->replaceImage($variableKey, $variableSrc, $drawingDOM, $options);

                                            if ($newRelsDrawing['imageFound']) {
                                                // add the new relationship
                                                $relsNodeDrawing = $drawingRelsDOM->createDocumentFragment();
                                                $relsNodeDrawing->appendXML($newRelsDrawing['rels']);
                                                $drawingRelsDOM->documentElement->appendChild($relsNodeDrawing);
                                            }
                                        }

                                        PhpxlsxLogger::logger('Replace variable image.', 'info');

                                        // refresh contents
                                        $this->zipXlsx->addContent($drawingTarget, $drawingDOM->saveXML());
                                        $this->zipXlsx->addContent($drawingTargetRels, $drawingRelsDOM->saveXML());

                                        // free DOMDocument resources
                                        $drawingDOM = null;
                                        $drawingRelsDOM = null;
                                    }
                                }
                            }

                            // free DOMDocument resources
                            $sheetRelsDOM = null;
                        }
                    }
                }

                // free DOMDocument resources
                $sheetDOM = null;
            }
        }
    }

    /**
     * Replaces an array of variables by their values
     *
     * @access public
     * @param array $variables variable names and new values
     * @param array $options
     *      'target' (string) sheets (default), headers, footers, comments
     */
    public function replaceVariableText($variables, $options = array())
    {
        if (!isset($options['target'])) {
            $options['target'] = 'sheets';
        }
        if (!isset($options['parseLineBreaks'])) {
            $options['parseLineBreaks'] = false;
        }

        PhpxlsxLogger::logger('Replace variable text.', 'info');

        if ($options['target'] == 'sheets') {
            $sharedStringsContents = $this->zipXlsx->getContentByType('sharedStrings');
            foreach ($sharedStringsContents as $sharedStringContents) {
                $newContent = $sharedStringContents['content'];
                foreach ($variables as $variableKey => $variableValue) {
                    $variableValue = $this->parseAndCleanTextString($variableValue);
                    $newContent = str_replace($this->templateSymbolStart . $variableKey . $this->templateSymbolEnd, $variableValue, $newContent);
                }

                $this->zipXlsx->addContent($sharedStringContents['path'], $newContent);
            }
        } else if ($options['target'] == 'headers' || $options['target'] == 'footers') {
            // get sheet contents to get headers and footers
            $sheetsContents = $this->zipXlsx->getSheets();
            foreach ($sheetsContents as $sheetContents) {
                // each sheet may have its own headers and footers. Get them
                $newContent = $sheetContents['content'];
                $sheetDOM = $this->zipXlsx->getContent($sheetContents['path'], 'DOMDocument');
                $sheetXPath = new \DOMXPath($sheetDOM);
                $sheetXPath->registerNamespace('xmlns', $this->namespaces['xmlns']);
                $queryContents = '';
                if ($options['target'] == 'headers') {
                    $queryContents = '//xmlns:headerFooter/xmlns:oddHeader|//xmlns:headerFooter/xmlns:evenHeader|//xmlns:headerFooter/xmlns:firstHeader';
                }
                if ($options['target'] == 'footers') {
                    $queryContents = '//xmlns:headerFooter/xmlns:oddFooter|//xmlns:headerFooter/xmlns:evenFooter|//xmlns:headerFooter/xmlns:firstFooter';
                }
                $headersFootersContents = $sheetXPath->query($queryContents);
                if ($headersFootersContents->length > 0) {
                    foreach ($headersFootersContents as $headersFootersContent) {
                        // keep header/footer init content. This variable will be used to replace the init contents by the new ones
                        $initContent = $headersFootersContent->ownerDocument->saveXML($headersFootersContent);
                        $newContentHeaderFooter = $initContent;
                        foreach ($variables as $variableKey => $variableValue) {
                            $variableValue = $this->parseAndCleanTextString($variableValue);
                            $newContentHeaderFooter = str_replace($this->templateSymbolStart . $variableKey . $this->templateSymbolEnd, $variableValue, $newContentHeaderFooter);
                        }
                        $newContent = str_replace($initContent, $newContentHeaderFooter, $newContent);
                    }

                    $this->zipXlsx->addContent($sheetContents['path'], $newContent);
                }

                // free DOMDocument resources
                $sheetDOM = null;
            }
        } else if ($options['target'] == 'comments') {
            $commentsContents = $this->zipXlsx->getContentByType('comments');
            foreach ($commentsContents as $commentContents) {
                $newContent = $commentContents['content'];
                foreach ($variables as $variableKey => $variableValue) {
                    $variableValue = $this->parseAndCleanTextString($variableValue);
                    $newContent = str_replace($this->templateSymbolStart . $variableKey . $this->templateSymbolEnd, $variableValue, $newContent);
                }

                $this->zipXlsx->addContent($commentContents['path'], $newContent);
            }
        }
    }

    /**
     * Extract the variables from a string
     *
     * @access private
     * @param string $content
     * @return array $variables
     */
    private function extractVariables($content) {
        $matches = array();
        preg_match_all('/'.preg_quote($this->templateSymbolStart, '/').self::$regExprVariableSymbols.preg_quote($this->templateSymbolEnd, '/').'/msiU', $content, $matches);

        $variables = array();
        foreach ($matches[0] as $variable) {
            $variables[] = str_replace(array($this->templateSymbolStart, $this->templateSymbolEnd), '', $variable);
        }

        return $variables;
    }

    /**
     * Replaces placeholder images
     *
     * @access public
     * @param string $variable this variable uniquely identifies the image we want to replace
     * @param string $src path, stream or base64 to the substitution image
     * @param \DOMDocument $domContent
     * @param string $options
     *      'target' (string) sheets, headers, footers
     * @throws \Exception image doesn't exist
     * @throws \Exception mime option is not set and getimagesizefromstring is not available
     * @throws \Exception mime option is not set and getimagesizefromstring is not available
     * @return string new rels
     */
    private function replaceImage($variable, $src, $domContent, $options = array())
    {
        // keep if the image has been found
        $imageFound = false;

        // get image information
        $imageInformation = new ImageUtilities();
        $imageContents = $imageInformation->returnImageContents($src, $options);

        // create a new Id
        $idImage = uniqid(rand(99,9999999));
        $ridImage = 'rId' . $idImage;
        // generate the new relationship
        $relString = '<Relationship Id="' . $ridImage . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/img' . $idImage . '.' . $imageContents['extension'] . '" />';
        // generate content type if it does not exist yet
        $this->generateDefault($imageContents['extension'], 'image/' . $imageContents['extension']);

        if ($options['target'] == 'sheets') {
            // transitional mode
            $domImages = $domContent->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'cNvPr');
            for ($i = 0; $i < $domImages->length; $i++) {
                if (
                        $domImages->item($i)->getAttribute('descr') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd ||
                        $domImages->item($i)->getAttribute('title') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd
                    ) {
                    // modify the image data to modify the r:embed attribute
                    $domImages->item($i)->parentNode->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')->item(0)->setAttribute('r:embed', $ridImage);

                    $imageFound = true;
                }
            }
            // strict mode
            $domImages = $domContent->getElementsByTagNameNS('http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing', 'cNvPr');
            for ($i = 0; $i < $domImages->length; $i++) {
                if (
                        $domImages->item($i)->getAttribute('descr') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd ||
                        $domImages->item($i)->getAttribute('title') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd
                    ) {
                    // modify the image data to modify the r:embed attribute
                    $domImages->item($i)->parentNode->parentNode->getElementsByTagNameNS('http://purl.oclc.org/ooxml/drawingml/main', 'blip')->item(0)->setAttribute('r:embed', $ridImage);

                    $imageFound = true;
                }
            }
        } else if ($options['target'] == 'headers' || $options['target'] == 'footers') {
            // shapes
            $domImages = $domContent->getElementsByTagNameNS('urn:schemas-microsoft-com:vml', 'shape');
            for ($i = 0; $i < $domImages->length; $i++) {
                if ($domImages->item($i)->getAttribute('alt') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd) {
                    // modify the image data to modify the o:relid attribute
                    // headers scope
                    if ($options['target'] == 'headers' && $domImages->item($i)->hasAttribute('id') && in_array($domImages->item($i)->getAttribute('id'), XlsxStructure::$idsHeaders)) {
                        $domImages->item($i)->getElementsByTagNameNS('urn:schemas-microsoft-com:vml', 'imagedata')->item(0)->setAttribute('o:relid', $ridImage);
                    }
                    // footers scope
                    if ($options['target'] == 'footers' && $domImages->item($i)->hasAttribute('id') && in_array($domImages->item($i)->getAttribute('id'), XlsxStructure::$idsFooters)) {
                        $domImages->item($i)->getElementsByTagNameNS('urn:schemas-microsoft-com:vml', 'imagedata')->item(0)->setAttribute('o:relid', $ridImage);
                    }

                    $imageFound = true;
                }
            }
            // drawing

            // transitional mode
            $domImages = $domContent->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'cNvPr');
            for ($i = 0; $i < $domImages->length; $i++) {
                if (
                    (
                        $domImages->item($i)->getAttribute('descr') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd ||
                        $domImages->item($i)->getAttribute('title') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd
                    )) {
                    // modify the image data to modify the r:embed attribute
                    // headers scope
                    if ($options['target'] == 'headers' && $domImages->item($i)->hasAttribute('name') && in_array($domImages->item($i)->getAttribute('name'), XlsxStructure::$idsHeaders)) {
                        $domImages->item($i)->parentNode->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')->item(0)->setAttribute('r:embed', $ridImage);
                    }
                    // footers scope
                    if ($options['target'] == 'footers' && $domImages->item($i)->hasAttribute('name') && in_array($domImages->item($i)->getAttribute('name'), XlsxStructure::$idsFooters)) {
                        $domImages->item($i)->parentNode->parentNode->getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')->item(0)->setAttribute('r:embed', $ridImage);
                    }

                    $imageFound = true;
                }
            }

            // strict mode
            $domImages = $domContent->getElementsByTagNameNS('http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing', 'cNvPr');
            for ($i = 0; $i < $domImages->length; $i++) {
                if (
                    (
                        $domImages->item($i)->getAttribute('descr') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd ||
                        $domImages->item($i)->getAttribute('title') == $this->templateSymbolStart . $variable . $this->templateSymbolEnd
                    )) {
                    // modify the image data to modify the r:embed attribute
                    // headers scope
                    if ($options['target'] == 'headers' && $domImages->item($i)->hasAttribute('name') && in_array($domImages->item($i)->getAttribute('name'), XlsxStructure::$idsHeaders)) {
                        $domImages->item($i)->parentNode->parentNode->getElementsByTagNameNS('http://purl.oclc.org/ooxml/drawingml/main', 'blip')->item(0)->setAttribute('r:embed', $ridImage);
                    }
                    // footers scope
                    if ($options['target'] == 'footers' && $domImages->item($i)->hasAttribute('name') && in_array($domImages->item($i)->getAttribute('name'), XlsxStructure::$idsFooters)) {
                        $domImages->item($i)->parentNode->parentNode->getElementsByTagNameNS('http://purl.oclc.org/ooxml/drawingml/main', 'blip')->item(0)->setAttribute('r:embed', $ridImage);
                    }

                    $imageFound = true;
                }
            }
        }

        if ($imageFound) {
            // copy the image in the template with the new name
            $this->zipXlsx->addContent('xl/media/img' . $idImage . '.' . $imageContents['extension'], $imageContents['content']);
        }

        return array('rels' => $relString, 'imageFound' => $imageFound);
    }
}