<?php

namespace Xls\Record;

class Externsheet extends AbstractRecord
{
    const NAME = 'EXTERNSHEET';
    const ID = 0x0017;

    /**
     * @param $refs
     *
     * @return string
     */
    public function getData($refs)
    {
        $refCount = count($refs);
        $data = pack('v', $refCount);

        foreach ($refs as $ref) {
            $data .= $ref;
        }

        return $this->getFullRecord($data);
    }
}
