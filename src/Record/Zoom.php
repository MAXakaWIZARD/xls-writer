<?php

namespace Xls\Record;

class Zoom extends AbstractRecord
{
    const NAME = 'ZOOM';
    const ID = 0x00A0;

    /**
     * Store the window zoom factor. This should be a reduced fraction but for
     * simplicity we will store all fractions with a numerator of 100.
     * @param $zoom
     *
     * @return string
     */
    public function getData($zoom)
    {
        $data = pack("vv", $zoom, 100);

        return $this->getFullRecord($data);
    }
}
