<?php

namespace Xls\Record;

class Country extends AbstractRecord
{
    const NAME = 'COUNTRY';
    const ID = 0x008C;
    const LENGTH = 0x04;

    /**
     * Write the COUNTRY record for localization
     * @param $countryCode
     *
     * @return string
     */
    public function getData($countryCode)
    {
        $data = pack('vv', $countryCode, $countryCode);

        return $this->getHeader() . $data;
    }
}
