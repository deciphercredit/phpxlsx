<?php
namespace Phpxlsx\Utilities;

use Phpxlsx\Logger\PhpxlsxLogger;
/**
 * Image functions
 *
 * @category   Phpxlsx
 * @package    utilities
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class ImageUtilities
{
    /**
     * Get image content and information
     *
     * @param string $image file path, base64, stream
     * @param array $options
     *      'mime' (string) forces a mime
     * @return array
     * @throws Exception image doesn't exist
     * @throws Exception image format is not supported
     * @throws Exception mime option is not set and getimagesizefromstring is not available
     */
    public function returnImageContents($image, $options = array())
    {
        $imageContents = array();

        if (strstr($image, 'base64,')) {
            // base64 image
            $descrArray = explode(';base64,', $image);
            $arrayExtension = explode('/', $descrArray[0]);
            $arrayMime = explode(':', $descrArray[0]);

            $imageContents['content'] = base64_decode($descrArray[1]);
            $imageContents['extension'] = strtolower($arrayExtension[1]);
            $imageContents['mime'] = $arrayMime[1];
            if (isset($options['mime'])) {
                $imageContents['mime'] = $options['mime'];
            }

            if (function_exists('getimagesizefromstring')) {
                // PHP 5.4 or newer
                $imageSize = getimagesizefromstring($imageContents['content']);
                $imageContents['width'] = $imageSize[0];
                $imageContents['height'] = $imageSize[1];
            }
        } else if (file_exists($image)) {
            // file content
            $extensionPath = pathinfo($image);

            $extension = strtolower($extensionPath['extension']);
            if (isset($options['mime'])) {
                $extension = $this->getExtensionFromMime($options['mime']);
            }

            $imageContents['content'] = file_get_contents($image);
            $imageContents['extension'] = $extension;
            $imageContents['mime'] = $this->getMimeFromExtension($extension);

            if (function_exists('getimagesizefromstring')) {
                // PHP 5.4 or newer
                $imageSize = getimagesizefromstring($imageContents['content']);
                $imageContents['width'] = $imageSize[0];
                $imageContents['height'] = $imageSize[1];
            } else {
                $imageSize = getimagesize($imageContents['content']);
                $imageContents['width'] = $imageSize[0];
                $imageContents['height'] = $imageSize[1];
            }
        } else {
            // stream content
            if (function_exists('getimagesizefromstring')) {
                $imageContents['content'] = file_get_contents($image);
                $attrImage = getimagesizefromstring($imageContents['content']);
                if (isset($options['mime'])) {
                    $attrImage['mime'] = $options['mime'];
                }
                $imageContents['extension'] = $this->getExtensionFromMime($attrImage['mime']);
                $imageContents['mime'] = $attrImage['mime'];
                $imageContents['width'] = $attrImage[0];
                $imageContents['height'] = $attrImage[1];
            } else {
                if (!isset($options['mime'])) {
                    PhpxlsxLogger::logger('getimagesizefromstring function is not available. Set the mime option.', 'fatal');
                }
            }
        }

        // check if the image can be obtained
        if (!isset($imageContents['content'])) {
            PhpxlsxLogger::logger('Unable to get the image.', 'fatal');
        }

        // check mime type
        if (!in_array($imageContents['mime'], array('image/jpg', 'image/jpeg', 'image/png', 'image/gif', 'image/bmp'))) {
            PhpxlsxLogger::logger('Image format \''.$imageContents['mime'].'\' is not supported.', 'fatal');
        }

        // rename jpg to jpeg, needed by Excel
        if ($imageContents['extension'] == 'jpg') {
            $imageContents['extension'] = 'jpeg';
        }

        return $imageContents;
    }


    /**
     * Gets extension from mime
     *
     * @access protected
     * @param string $mime
     * @return string
     */
    protected function getExtensionFromMime($mime)
    {
        $extension = '';

        switch ($mime) {
            case 'image/bmp':
                $extension = 'bmp';
                break;
            case 'image/gif':
                $extension = 'gif';
                break;
            case 'image/jpg':
            case 'image/jpeg':
                $extension = 'jpeg';
                break;
            case 'image/png':
                $extension = 'png';
                break;
            default:
                break;
        }

        return strtolower($extension);
    }

    /**
     * Gets mime from extension
     *
     * @access protected
     * @param string $extension
     * @return string
     */
    protected function getMimeFromExtension($extension)
    {
        $mime = '';

        switch ($extension) {
            case 'bmp':
                $mime = 'image/bmp';
                break;
            case 'gif':
                $mime = 'image/gif';
                break;
            case 'jpg':
                $mime = 'image/jpg';
                break;
            case 'jpeg':
                $mime = 'image/jpeg';
                break;
            case 'png':
                $mime = 'image/png';
                break;
            default:
                break;
        }

        return strtolower($mime);
    }
}