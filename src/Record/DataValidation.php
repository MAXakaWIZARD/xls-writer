<?php

namespace Xls\Record;

class DataValidation extends AbstractRecord
{
    const NAME = 'DATAVALIDATION';
    const ID = 0x01BE;

    /**
     * Generate the DVAL biff record
     * @param $dv
     *
     * @return string
     */
    public function getData($dv)
    {
        return $this->getFullRecord($dv);
    }
}
