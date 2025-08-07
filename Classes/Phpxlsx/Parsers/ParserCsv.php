<?php
namespace Phpxlsx\Parsers;

use Phpxlsx\Logger\PhpxlsxLogger;

/**
 * Parser CSV
 *
 * @category   Phpxlsx
 * @package    parsers
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class ParserCsv
{
    /**
     * Parser
     *
     * @access public
     * @param string $csv CSV path
     * @param array $options
     *      'delimiter' (string) field separator. Default as ,
     *      'enclosure' (string) field enclosure character. Default as "
     *      'escape' (string) escape character. Default as \
     *      'firstRowAsHeader' (bool) if true, the first row is added as table header. Default as false
     * @throws Exception CSV format is not supported
     * @return array
     */
    public function parser($csv, $options)
    {
        $content = array();
        if (($handleFile = fopen($csv, 'r')) !== FALSE) {
            while (($row = fgetcsv($handleFile, null, $options['delimiter'], $options['enclosure'], $options['escape'])) !== FALSE) {
                $content[] = $row;
            }
            fclose($handleFile);
        } else {
            PhpxlsxLogger::logger('Unable to get the CSV.', 'fatal');
        }

        return $content;
    }
}