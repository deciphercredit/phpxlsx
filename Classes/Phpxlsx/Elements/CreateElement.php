<?php
namespace Phpxlsx\Elements;
/**
 * Create tag elements
 *
 * @category   Phpxlsx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class CreateElement
{
    /**
     * Parse and clean a text string to be added
     *
     * @access public
     * @var string $content
     * @var string $content
     */
    public function parseAndCleanTextString($content) {
        $content = htmlspecialchars($content);

        return $content;
    }

    /**
     * Transform a string position to int starting from 0
     *
     * @access public
     * @param string $value
     * @return int
     */
    public function wordToInt($value) {
        $valueInt = 0;

        $strValue = array_reverse(str_split($value));

        for($i = 0; $i < strlen($value); $i++) {
            $valueInt += (ord($strValue[$i])-64) * pow(26,$i);
        }
        // 0 is the first value, not 1
        $valueInt--;

        return $valueInt;
    }
}