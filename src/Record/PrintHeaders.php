<?php
namespace Xls\Record;

class PrintHeaders extends AbstractRecord
{
    const NAME = 'PRINTHEADERS';
    const ID = 0x2A;
    const LENGTH = 0x02;

    /**
     * @param $printRowColHeaders
     *
     * @return string
     */
    public function getData($printRowColHeaders)
    {
        $data = pack("v", intval($printRowColHeaders));

        return $this->getHeader() . $data;
    }
}
