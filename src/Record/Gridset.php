<?php
namespace Xls\Record;

class Gridset extends AbstractRecord
{
    const NAME = 'GRIDSET';
    const ID = 0x82;

    /**
     * @param $gridsetVisible
     *
     * @return string
     */
    public function getData($gridsetVisible)
    {
        $data = pack("v", intval($gridsetVisible));

        return $this->getFullRecord($data);
    }
}
