<?php
namespace Xls;

class StringUtils
{
    const STRING_REGEXP_FRACTION = '(-?)(\d+)\s+(\d+\/\d+)';

    /**
     * Get whether mbstring extension is available
     *
     * @return boolean
     */
    public static function isMbstringEnabled()
    {
        return function_exists('mb_convert_encoding');
    }

    /**
     * Get whether iconv extension is available
     *
     * @return boolean
     */
    public static function isIconvEnabled()
    {
        return function_exists('iconv');
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
        return self::toBiff8Unicode($value, 8);
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
        return self::toBiff8Unicode($value, 16);
    }

    /**
     * @param string $value
     * @param int $lengthSize 8 or 16
     *
     * @return string
     */
    public static function toBiff8Unicode($value, $lengthSize)
    {
        $ln = self::countCharacters($value);
        $lengthFormat = ($lengthSize == 8) ? 'C' : 'v';
        $data = pack($lengthFormat, $ln);

        $opt = (self::isMbstringOrIconvEnabled()) ? 0x01 : 0x00;
        $data .= pack('C', $opt);

        $data .= self::toUtf16Le($value);

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
     * @param string $from Encoding to convert from, e.g. 'UTF-16LE'
     * @param string $to Encoding to convert to, e.g. 'UTF-8'
     * @return string
     */
    public static function convertEncoding($value, $from, $to)
    {
        if (self::isMbstringEnabled()) {
            return mb_convert_encoding($value, $to, $from);
        }

        if (self::isIconvEnabled()) {
            return iconv($from, $to, $value);
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
        return self::convertEncoding($value, 'UTF-8', 'UTF-16LE');
    }

    /**
     * @param $str
     *
     * @return string
     */
    public static function toNullTerminatedWchar($str)
    {
        $str = join("\0", preg_split("''", $str, -1, PREG_SPLIT_NO_EMPTY));
        $str = $str . "\0\0\0";

        return $str;
    }
}
