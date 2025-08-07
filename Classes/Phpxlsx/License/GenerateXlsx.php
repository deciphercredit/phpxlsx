<?php
namespace Phpxlsx\License;
/**
 * Check for a valid license
 *
 * @category   Phpxlsx
 * @package    license
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class GenerateXlsx
{
    /**
     * Check for a valid license
     *
     * @access public
     * @return boolean
     */
    public static function beginXlsx()
    {
        $xzerod = '';
        $xzeroc = '';
        $xzeroi = '';
        $phpxlsxconfig = \Phpxlsx\Utilities\PhpxlsxUtilities::parseConfig();

        if (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off') {
            return;
        }

        if (!isset($_SERVER['SERVER_NAME']) || !isset($_SERVER['SERVER_ADDR'])) {
            return;
        } else {
            $xzerod = trim($phpxlsxconfig['license']['code']);
            $xzeroc = trim($_SERVER['SERVER_NAME']);
            $xzeroi = trim($_SERVER['SERVER_ADDR']);

            if (empty($xzeroi)) {
                $xzeroi = $xzeroc;
            }
        }
        if (
            preg_match('/^192.168./', $xzeroi) ||
            preg_match('/^172./', $xzeroi) ||
            preg_match('/^10./', $xzeroi) ||
            preg_match('/^127./', $xzeroi) ||
            preg_match('/^::1/', $xzeroi) ||
            preg_match('/localhost/', $xzeroc)
        ) {
            return;
        } elseif ($xzerod == md5($xzeroc . '_basic_xlsx')) {
            return;
        } elseif ($xzerod == md5($xzeroc . '_advanced_xlsx')) {
            return;
        } elseif ($xzerod == md5($xzeroc . '_premium_xlsx')) {
            return;
        } elseif ($xzerod == md5($xzeroi . '_premium_xlsx')) {
            return;
        }

        if (!preg_match('/^www./', $xzeroc)) {
            $xzeroc = 'www.' . $xzeroc;
        }
        if ($xzerod == md5($xzeroc . '_basic_xlsx')) {
            return;
        } elseif ($xzerod == md5($xzeroc . '_advanced_xlsx')) {
            return;
        } elseif ($xzerod == md5($xzeroc . '_premium_xlsx')) {
            return;
        }

        $serverNameSeg = explode('.', trim($_SERVER['SERVER_NAME']));
        $serverNamePart = '';
        $serverNameSegI = count($serverNameSeg);
        for ($i = $serverNameSegI-1; $i >= 0; $i--) {
            if (empty($serverNamePart)) {
                $serverNamePart = $serverNameSeg[$i];
            } else {
                $serverNamePart = $serverNameSeg[$i] . '.' . $serverNamePart;
            }
            if ($xzerod == md5($serverNamePart . '_basic_xlsx')) {
                return;
            } elseif ($xzerod == md5($serverNamePart . '_advanced_xlsx')) {
                return;
            } elseif ($xzerod == md5($serverNamePart . '_premium_xlsx')) {
                return;
            }
        }

        throw new \Exception('There is not a valid license');
    }

}
