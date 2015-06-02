<?php
namespace Xls\Record;

class RecalcId extends AbstractRecord
{
    const NAME = 'RECALCID';
    const ID = 0x01C1;

    /**
     *
     * @return string
     */
    public function getData()
    {
        $data = pack('VV', 0x000001C1, 0x00001E667);

        return $this->getFullRecord($data);
    }
}
