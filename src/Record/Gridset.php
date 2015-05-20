<?php
namespace Xls\Record;

class Gridset extends AbstractRecord
{
    const NAME = 'GRIDSET';
    const ID = 0x82;
    const LENGTH = 0x02;

    /**
     * @param $gridsetVisible
     *
     * @return string
     */
    public function getData($gridsetVisible)
    {
        $data = pack("v", intval($gridsetVisible));

        return $this->getHeader() . $data;
    }
}
