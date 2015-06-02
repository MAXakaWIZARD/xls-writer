<?php

namespace Xls\Record;

class BitmapCoreHeader extends AbstractRecord
{
    const NAME = 'BITMAPCOREHEADER';

    /**
     * @param $width
     * @param $height
     * @return string
     */
    public function getData($width, $height)
    {
        $planesCount = 1;
        $colorDepth = 24;

        $data = pack("vvvv", $width, $height, $planesCount, $colorDepth);

        return pack("V", strlen($data) + 4) . $data;
    }
}
