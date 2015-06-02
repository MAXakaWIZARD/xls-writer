<?php

namespace Xls\Record;

class Defcolwidth extends AbstractRecord
{
    const NAME = 'DEFCOLWIDTH';
    const ID = 0x0055;

    /**
     * Generate the DEFCOLWIDTH record
     *
     * @return string
     */
    public function getData()
    {
        $colwidth = 0x0008; // Default column width
        $data = pack("v", $colwidth);

        return $this->getFullRecord($data);
    }
}
