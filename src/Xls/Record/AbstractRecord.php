<?php

namespace Xls\Record;

abstract class AbstractRecord
{
    /**
     * Returns record header data ready for writing
     * @param int $extraLength
     * @param int $extraParam
     * @return string
     */
    protected function getHeader($extraLength = 0, $extraParam = null)
    {
        $length = static::LENGTH + $extraLength;

        if (is_null($extraParam)) {
            return pack("vv", static::ID, $length);
        } else {
            return pack("vvv", static::ID, $length, $extraParam);
        }
    }
}
