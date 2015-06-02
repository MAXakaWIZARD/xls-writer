<?php

namespace Xls\Record;

class Country extends AbstractRecord
{
    const NAME = 'COUNTRY';
    const ID = 0x008C;

    /**
     * Generate the COUNTRY record for localization
     * @param $countryCode
     *
     * @return string
     */
    public function getData($countryCode)
    {
        $data = pack('vv', $countryCode, $countryCode);

        return $this->getFullRecord($data);
    }
}
