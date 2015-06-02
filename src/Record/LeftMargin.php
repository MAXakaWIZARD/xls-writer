<?php

namespace Xls\Record;

use Xls\BIFFwriter;

class LeftMargin extends AbstractRecord
{
    const NAME = 'LEFTMARGIN';
    const ID = 0x26;

    /**
     * @param $margin
     *
     * @return string
     */
    public function getData($margin)
    {
        $data = pack("d", $margin);
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
            $data = strrev($data);
        }

        return $this->getFullRecord($data);
    }
}
