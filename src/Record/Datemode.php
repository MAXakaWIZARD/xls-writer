<?php

namespace Xls\Record;

class Datemode extends AbstractRecord
{
    const NAME = 'DATEMODE';
    const ID = 0x0022;

    /**
     * Generate DATEMODE record to indicate the date system in use (1904 or 1900).
     *
     * @param integer $f1904 Flag for 1904 date system (0 => base date is 1900, 1 => base date is 1904)
     * @return string
     */
    public function getData($f1904)
    {
        $data = pack("v", $f1904);

        return $this->getFullRecord($data);
    }
}
