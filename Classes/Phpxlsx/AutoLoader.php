<?php
namespace Phpxlsx;
/**
 * Autoloader
 *
 * @category   Phpxlsx
 * @package    loader
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpxlsx LICENSE
 * @link       https://www.phpxlsx.com
 */
class AutoLoader
{
    /**
     * Main tags of relationships XML
     *
     * @access public
     * @static
     */
    public static function load()
    {
        spl_autoload_register(array('Phpxlsx\AutoLoader', 'autoloadGenericClasses'));
    }

    /**
     * Autoload phpxlsx
     *
     * @access public
     * @param string $className Class to load
     */
    public static function autoloadGenericClasses($namespace)
    {
        $splitpath = explode('\\', $namespace);
        $path = '';
        $name = '';
        $firstword = true;
        for ($i = 0; $i < count($splitpath); $i++) {
            if ($splitpath[$i] && !$firstword) {
                if ($i == count($splitpath) - 1){
                    $name = $splitpath[$i];
                }
                else{
                    $path .= DIRECTORY_SEPARATOR . $splitpath[$i];
                }
            }
            if ($splitpath[$i] && $firstword) {
                if ($splitpath[$i] != __NAMESPACE__){
                    break;
                }
                $firstword = false;
            }
        }
        if (!$firstword) {
            $fullpath = __DIR__ . $path . DIRECTORY_SEPARATOR . $name . '.php';
            if (file_exists($fullpath)) {
                return include_once($fullpath);
            }

        }
        return false;
    }

}
