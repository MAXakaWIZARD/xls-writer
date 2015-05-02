<?php

namespace Xls\Record;

class Hcenter extends AbstractRecord
{
    const NAME = 'HCENTER';
    const ID = 0x0083;
    const LENGTH = 0x02;

    /**
     * @param $centering
     *
     * @return string
     */
    public function getData($centering)
    {
        $data = pack('v', $centering);

        return $this->getHeader() . $data;
    }
}
