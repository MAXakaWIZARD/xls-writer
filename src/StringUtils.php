<?php
namespace Xls;

class StringUtils
{
    const STRING_REGEXP_FRACTION = '(-?)(\d+)\s+(\d+\/\d+)';

    /**
     * Is mbstring extension avalable?
     *
     * @var boolean
     */
    private static $isMbstringEnabled;

    /**
     * Is iconv extension avalable?
     *
     * @var boolean
     */
    private static $isIconvEnabled;

    /**
     * Get whether mbstring extension is available
     *
     * @return boolean
     */
    public static function isMbstringEnabled()
    {
        if (isset(self::$isMbstringEnabled)) {
            return self::$isMbstringEnabled;
        }

        self::$isMbstringEnabled = function_exists('mb_convert_encoding');

        return self::$isMbstringEnabled;
    }

    /**
     * Get whether iconv extension is available
     *
     * @return boolean
     */
    public static function isIconvEnabled()
    {
        if (isset(self::$isIconvEnabled)) {
            return self::$isIconvEnabled;
        }

        // Fail if iconv doesn't exist
        if (!function_exists('iconv')) {
            self::$isIconvEnabled = false;
            return false;
        }

        // Sometimes iconv is not working, and e.g. iconv('UTF-8', 'UTF-16LE', 'x') just returns false,
        if (!@iconv('UTF-8', 'UTF-16LE', 'x')) {
            self::$isIconvEnabled = false;
            return false;
        }

        // Sometimes iconv_substr('A', 0, 1, 'UTF-8') just returns false in PHP 5.2.0
        // we cannot use iconv in that case either (http://bugs.php.net/bug.php?id=37773)
        if (!@iconv_substr('A', 0, 1, 'UTF-8')) {
            self::$isIconvEnabled = false;
            return false;
        }

        // CUSTOM: IBM AIX iconv() does not work
        if (defined('PHP_OS')
            && @stristr(PHP_OS, 'AIX')
            && defined('ICONV_IMPL')
            && (@strcasecmp(ICONV_IMPL, 'unknown') == 0)
            && defined('ICONV_VERSION')
            && (@strcasecmp(ICONV_VERSION, 'unknown') == 0)
        ) {
            self::$isIconvEnabled = false;
            return false;
        }

        // If we reach here no problems were detected with iconv
        self::$isIconvEnabled = true;

        return true;
    }

    /**
     * Try to sanitize UTF8, stripping invalid byte sequences. Not perfect. Does not surrogate characters.
     *
     * @param string $value
     * @return string
     */
    public static function sanitizeUTF8($value)
    {
        if (self::isIconvEnabled()) {
            $value = @iconv('UTF-8', 'UTF-8', $value);
            return $value;
        }

        if (self::isMbstringEnabled()) {
            $value = mb_convert_encoding($value, 'UTF-8', 'UTF-8');
            return $value;
        }

        return $value;
    }

    /**
     * Check if a string contains UTF8 data
     *
     * @param string $value
     * @return boolean
     */
    public static function isUTF8($value = '')
    {
        return $value === '' || preg_match('/^./su', $value) === 1;
    }

    /**
     * Formats a numeric value as a string for output in various output writers forcing
     * point as decimal separator in case locale is other than English.
     *
     * @param mixed $value
     * @return string
     */
    public static function formatNumber($value)
    {
        if (is_float($value)) {
            return str_replace(',', '.', $value);
        }

        return (string) $value;
    }

    /**
     * Converts a UTF-8 string into BIFF8 Unicode string data (8-bit string length)
     * Writes the string using uncompressed notation, no rich text, no Asian phonetics
     * If mbstring extension is not available, ASCII is assumed, and compressed notation is used
     * although this will give wrong results for non-ASCII strings
     * see OpenOffice.org's Documentation of the Microsoft Excel File Format, sect. 2.5.3
     *
     * @param string  $value    UTF-8 encoded string
     * @param mixed[] $arrcRuns Details of rich text runs in $value
     * @return string
     */
    public static function UTF8toBIFF8UnicodeShort($value, $arrcRuns = array())
    {
        // character count
        $ln = self::CountCharacters($value, 'UTF-8');
        // option flags
        if (empty($arrcRuns)) {
            $opt = (self::isIconvEnabled() || self::isMbstringEnabled()) ?
                0x0001 : 0x0000;
            $data = pack('CC', $ln, $opt);
            // characters
            $data .= self::ConvertEncoding($value, 'UTF-16LE', 'UTF-8');
        } else {
            $data = pack('vC', $ln, 0x09);
            $data .= pack('v', count($arrcRuns));
            // characters
            $data .= self::ConvertEncoding($value, 'UTF-16LE', 'UTF-8');
            foreach ($arrcRuns as $cRun) {
                $data .= pack('v', $cRun['strlen']);
                $data .= pack('v', $cRun['fontidx']);
            }
        }

        return $data;
    }

    /**
     * Converts a UTF-8 string into BIFF8 Unicode string data (16-bit string length)
     * Writes the string using uncompressed notation, no rich text, no Asian phonetics
     * If mbstring extension is not available, ASCII is assumed, and compressed notation is used
     * although this will give wrong results for non-ASCII strings
     * see OpenOffice.org's Documentation of the Microsoft Excel File Format, sect. 2.5.3
     *
     * @param string $value UTF-8 encoded string
     * @return string
     */
    public static function UTF8toBIFF8UnicodeLong($value)
    {
        // character count
        $ln = self::CountCharacters($value, 'UTF-8');

        // option flags
        $opt = (self::isIconvEnabled() || self::isMbstringEnabled()) ?
            0x0001 : 0x0000;

        // characters
        $chars = self::ConvertEncoding($value, 'UTF-16LE', 'UTF-8');

        $data = pack('vC', $ln, $opt) . $chars;

        return $data;
    }

    /**
     * Convert string from one encoding to another. First try mbstring, then iconv, finally strlen
     *
     * @param string $value
     * @param string $to Encoding to convert to, e.g. 'UTF-8'
     * @param string $from Encoding to convert from, e.g. 'UTF-16LE'
     * @return string
     */
    public static function convertEncoding($value, $to, $from)
    {
        if (self::isIconvEnabled()) {
            return iconv($from, $to, $value);
        }

        if (self::isMbstringEnabled()) {
            return mb_convert_encoding($value, $to, $from);
        }

        if ($from == 'UTF-16LE' || $from == 'UTF-16BE') {
            $isBigEndian = $from == 'UTF-16BE';
            return self::utf16Decode($value, $isBigEndian);
        }

        return $value;
    }

    /**
     * Decode UTF-16 encoded strings.
     *
     * Can handle both BOM'ed data and un-BOM'ed data.
     * Assumes Big-Endian byte order if no BOM is available.
     * This function was taken from http://php.net/manual/en/function.utf8-decode.php
     * and $bom_be parameter added.
     *
     * @param   string  $str  UTF-16 encoded data to decode.
     * @param bool $bom_be
     * @return  string  UTF-8 / ISO encoded data.
     */
    public static function utf16Decode($str, $bom_be = true)
    {
        if (strlen($str) < 2) {
            return $str;
        }

        $c0 = ord($str{0});
        $c1 = ord($str{1});
        $str = substr($str, 2);

        if ($c0 == 0xff && $c1 == 0xfe) {
            $bom_be = false;
        }

        $len = strlen($str);
        $newstr = '';
        for ($i = 0; $i < $len; $i += 2) {
            if ($bom_be) {
                $val = ord($str{$i}) << 4;
                $val += ord($str{$i+1});
            } else {
                $val = ord($str{$i+1}) << 4;
                $val += ord($str{$i});
            }

            $newstr .= ($val == 0x228) ? "\n" : chr($val);
        }

        return $newstr;
    }

    /**
     * Get character count. First try mbstring, then iconv, finally strlen
     *
     * @param string $value
     * @param string $enc Encoding
     * @return int Character count
     */
    public static function countCharacters($value, $enc = 'UTF-8')
    {
        if (self::isMbstringEnabled()) {
            return mb_strlen($value, $enc);
        }

        if (self::isIconvEnabled()) {
            return iconv_strlen($value, $enc);
        }

        // else strlen
        return strlen($value);
    }

    /**
     * Get a substring of a UTF-8 encoded string. First try mbstring, then iconv, finally strlen
     *
     * @param string $pValue UTF-8 encoded string
     * @param int $pStart Start offset
     * @param int $pLength Maximum number of characters in substring
     * @return string
     */
    public static function substring($pValue = '', $pStart = 0, $pLength = 0)
    {
        if (self::isMbstringEnabled()) {
            return mb_substr($pValue, $pStart, $pLength, 'UTF-8');
        }

        if (self::isIconvEnabled()) {
            return iconv_substr($pValue, $pStart, $pLength, 'UTF-8');
        }

        // else substr
        return substr($pValue, $pStart, $pLength);
    }

    /**
     * Convert a UTF-8 encoded string to upper case
     *
     * @param string $pValue UTF-8 encoded string
     * @return string
     */
    public static function strToUpper($pValue = '')
    {
        if (function_exists('mb_convert_case')) {
            return mb_convert_case($pValue, MB_CASE_UPPER, "UTF-8");
        }

        return strtoupper($pValue);
    }

    /**
     * Convert a UTF-8 encoded string to lower case
     *
     * @param string $pValue UTF-8 encoded string
     * @return string
     */
    public static function strToLower($pValue = '')
    {
        if (function_exists('mb_convert_case')) {
            return mb_convert_case($pValue, MB_CASE_LOWER, "UTF-8");
        }

        return strtolower($pValue);
    }

    /**
     * Convert a UTF-8 encoded string to title/proper case
     *    (uppercase every first character in each word, lower case all other characters)
     *
     * @param string $pValue UTF-8 encoded string
     * @return string
     */
    public static function strToTitle($pValue = '')
    {
        if (function_exists('mb_convert_case')) {
            return mb_convert_case($pValue, MB_CASE_TITLE, "UTF-8");
        }

        return ucwords($pValue);
    }

    public static function mbIsUpper($char)
    {
        return mb_strtolower($char, "UTF-8") != $char;
    }

    public static function mbStrSplit($string)
    {
        # Split at all position not after the start: ^
        # and not before the end: $
        return preg_split('/(?<!^)(?!$)/u', $string);
    }

    /**
     * Reverse the case of a string, so that all uppercase characters become lowercase
     *    and all lowercase characters become uppercase
     *
     * @param string $pValue UTF-8 encoded string
     * @return string
     */
    public static function strCaseReverse($pValue = '')
    {
        if (self::isMbstringEnabled()) {
            $characters = self::mbStrSplit($pValue);
            foreach ($characters as &$character) {
                if (self::mbIsUpper($character)) {
                    $character = mb_strtolower($character, 'UTF-8');
                } else {
                    $character = mb_strtoupper($character, 'UTF-8');
                }
            }

            return implode('', $characters);
        }

        return strtolower($pValue) ^ strtoupper($pValue) ^ $pValue;
    }

    /**
     * Retrieve any leading numeric part of a string, or return the full string if no leading numeric
     *    (handles basic integer or float, but not exponent or non decimal)
     *
     * @param    string    $value
     * @return    mixed    string or only the leading numeric part of the string
     */
    public static function testStringAsNumeric($value)
    {
        if (is_numeric($value)) {
            return $value;
        }

        $v = floatval($value);

        return (is_numeric(substr($value, 0, strlen($v)))) ? $v : $value;
    }

    /**
     * @param $value
     *
     * @return string
     */
    public static function toUtf16Le($value)
    {
        if (!self::isIconvEnabled() || !self::isMbstringEnabled()) {
            return $value;
        }

        $encoding = mb_detect_encoding($value, 'auto');
        if ($encoding !== 'UTF-16LE') {
            $value = iconv($encoding, 'UTF-16LE', $value);
        }

        return $value;
    }

    /**
     * @param $encoding
     *
     * @throws \Exception
     */
    public static function checkEncoding($encoding)
    {
        if ($encoding != 'UTF-16LE' && !self::isIconvEnabled()) {
            throw new \Exception("Using an input encoding other than UTF-16LE requires PHP support for iconv");
        }
    }
}
