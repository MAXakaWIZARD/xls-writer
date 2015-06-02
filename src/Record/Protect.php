<?php

namespace Xls\Record;

class Protect extends AbstractRecord
{
    const NAME = 'PROTECT';
    const ID = 0x0012;

    /**
     * Generate the PROTECT biff record
     *
     * @param integer $lock
     *
     * @return string
     */
    public function getData($lock)
    {
        $data = pack("v", $lock);

        return $this->getFullRecord($data);
    }
}
