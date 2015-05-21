<?php

namespace Xls\Record;

class ExternalBook extends AbstractRecord
{
    const NAME = 'EXTERNALBOOK';
    const ID = 0x01AE;
    const LENGTH = 0x04;

    /**
     * Generate Internal SUPBOOK record
     * @param int $worksheetsCount
     *
     * @return string
     */
    public function getData($worksheetsCount)
    {
        $data = pack("vv", $worksheetsCount, 0x0401);

        return $this->getHeader() . $data;
    }
}
