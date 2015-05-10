<?php

namespace Xls\Record;

class BitmapCoreHeader extends AbstractRecord
{
    const NAME = 'BITMAPCOREHEADER';
    const ID = 0x00;
    const LENGTH = 0x0C;

    /**
     * Generate the BITMAPCOREHEADER biff record
     * @param $width
     * @param $height
     * @param $data
     *
     * @return string
     */
    public function getData($width, $height, $data)
    {
        $header = pack("Vvvvv", self::LENGTH, $width, $height, 0x01, 0x18);

        return $header . $data;
    }
}
