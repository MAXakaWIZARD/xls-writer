<?php

namespace Xls\OLE;

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
            return str_repeat("\x00", 8);
        }

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
        $bigDate *= 10000000;

        // factor used for separating numbers into 4 bytes parts
        $factor = pow(2, 32);

        $div = $bigDate / $factor;
        $highPart = floor($div);
        // lower 4 bytes
        $lowPart = floor(($div - $highPart) * $factor);

        return pack('V2', $lowPart, $highPart);
    }
}
