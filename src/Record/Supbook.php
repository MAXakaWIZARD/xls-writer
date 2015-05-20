<?php

namespace Xls\Record;

class Supbook extends AbstractRecord
{
    const NAME = 'SUPBOOK';
    const ID = 0x01AE;
    const LENGTH = 0x04;

    /**
     * Generate Internal SUPBOOK record
     * @param $worksheets
     *
     * @return string
     */
    public function getData($worksheets)
    {
        $data = pack("vv", count($worksheets), 0x0104);

        return $this->getHeader() . $data;
    }
}
