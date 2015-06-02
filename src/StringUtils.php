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
        if (!iconv('UTF-8', 'UTF-16LE', 'x')) {
            self::$isIconvEnabled = false;

            return false;
        }

        // If we reach here no problems were detected with iconv
        self::$isIconvEnabled = true;

        return true;
    }

    /**
     * @return bool
     */
    public static function isMbstringOrIconvEnabled()
    {
        return self::isMbstringEnabled() || self::isIconvEnabled();
    }

    /**
     * Converts a UTF-8 string into BIFF8 Unicode string data (8-bit string length)
     * Writes the string using uncompressed notation, no rich text, no Asian phonetics
     * If mbstring extension is not available, ASCII is assumed, and compressed notation is used
     * although this will give wrong results for non-ASCII strings
     * see OpenOffice.org's Documentation of the Microsoft Excel File Format, sect. 2.5.3
     *
     * @param string  $value    UTF-8 encoded string
     * @return string
     */
    public static function toBiff8UnicodeShort($value)
    {
        $ln = self::CountCharacters($value, 'UTF-8');
        $opt = (self::isMbstringOrIconvEnabled()) ? 0x01 : 0x00;
        $data = pack('CC', $ln, $opt);
        $data .= self::ConvertEncoding($value, 'UTF-16LE', 'UTF-8');

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
    public static function toBiff8UnicodeLong($value)
    {
        // character count
        $ln = self::CountCharacters($value, 'UTF-8');

        // option flags
        $opt = (self::isMbstringOrIconvEnabled()) ? 0x01 : 0x00;

        // characters
        $chars = self::toUtf16Le($value);

        $data = pack('vC', $ln, $opt) . $chars;

        return $data;
    }

    /**
     * @param $value
     *
     * @return string
     */
    public static function toBiff8UnicodeLongWoLenInfo($value)
    {
        return substr(self::toBiff8UnicodeLong($value), 2);
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

        return $value;
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
     * @param string $str UTF-8 encoded string
     * @param int $start Start offset
     * @param int $length Maximum number of characters in substring
     *
     * @return string
     */
    public static function substr($str = '', $start = 0, $length = 0)
    {
        if (self::isMbstringEnabled()) {
            return mb_substr($str, $start, $length, 'UTF-8');
        }

        if (self::isIconvEnabled()) {
            return iconv_substr($str, $start, $length, 'UTF-8');
        }

        // else substr
        return substr($str, $start, $length);
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
            $value = self::convertEncoding($value, 'UTF-16LE', $encoding);
        }

        return $value;
    }

    public static function toNullTerminatedWchar($str)
    {
        $str = join("\0", preg_split("''", $str, -1, PREG_SPLIT_NO_EMPTY));
        $str = $str . "\0\0\0";

        return $str;
    }
}
