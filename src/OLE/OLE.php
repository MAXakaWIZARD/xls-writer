<?php

namespace Xls\OLE;

/**
 * OLE package base class.
 *
 * @category Structures
 * @package  OLE
 * @author   Xavier Noguer <xnoguer@php.net>
 * @author   Christian Schmidt <schmidt@php.net>
 */
class OLE
{
    /**
     * Utility function to transform ASCII text to Unicode
     *
     * @param string $ascii The ASCII string to transform
     * @return string The string in Unicode
     */
    public static function asc2Ucs($ascii)
    {
        $rawname = '';
        $len = strlen($ascii);
        for ($i = 0; $i < $len; $i++) {
            $rawname .= $ascii{$i} . "\x00";
        }

        return $rawname;
    }

    /**
     * Returns a string for the OLE container with the date given
     *
     * @param integer $date A timestamp
     *
     * @return string The string for the OLE container
     */
    public static function localDate2OLE($date = null)
    {
        if (!isset($date)) {
            return "\x00\x00\x00\x00\x00\x00\x00\x00";
        }

        // factor used for separating numbers into 4 bytes parts
        $factor = pow(2, 32);

        // days from 1-1-1601 until the beggining of UNIX era
        $days = 134774;
        // calculate seconds
        $bigDate = $days * 24 * 3600 +
            gmmktime(
                date("H", $date),
                date("i", $date),
                date("s", $date),
                date("m", $date),
                date("d", $date),
                date("Y", $date)
            );
        // multiply just to make MS happy
        $bigDate *= 10000000;

        $highPart = floor($bigDate / $factor);
        // lower 4 bytes
        $lowPart = floor((($bigDate / $factor) - $highPart) * $factor);

        // Make HEX string
        $res = '';

        for ($i = 0; $i < 4; $i++) {
            $hex = $lowPart % 0x100;
            $res .= pack('c', $hex);
            $lowPart /= 0x100;
        }

        for ($i = 0; $i < 4; $i++) {
            $hex = $highPart % 0x100;
            $res .= pack('c', $hex);
            $highPart /= 0x100;
        }

        return $res;
    }
}
