<?php

namespace Xls\Record;

class Dv extends AbstractRecord
{
    const NAME = 'DV';
    const ID = 0x01BE;

    /**
     * Generate the DVAL biff record
     * @param $dv
     *
     * @return string
     */
    public function getData($dv)
    {
        $extraLength = strlen($dv);

        return $this->getHeader($extraLength) . $dv;
    }
}
