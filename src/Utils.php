<?php

namespace Xls;

class Utils
{
    /**
     * returns hex representation of binary data
     * @param $data
     *
     * @return string
     */
    public static function hexDump($data)
    {
        $result = '';

        $charCount = strlen($data);
        for ($i = 0; $i < $charCount; $i++) {
            $byte = ord($data[$i]);
            if ($i > 0) {
                $result .= ' ';
            }
            $result .= sprintf('%02X', $byte);
        }

        return $result;
    }

    /**
     * @return string
     */
    public static function generateGuid()
    {
        return strtoupper(md5(uniqid(rand(), true)));
    }
}
