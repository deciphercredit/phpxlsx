<?php
namespace Phpxlsx\Sign;
/**
 * Sign a file
 *
 * @category   Phpxlsx
 * @package    sign
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
interface Sign
{
    /**
     * Setter $_privatekey
     */
    public function setPrivateKey($file, $password = null);

    /**
     * Setter $_x509Certificate
     */
    public function setX509Certificate($file);
}