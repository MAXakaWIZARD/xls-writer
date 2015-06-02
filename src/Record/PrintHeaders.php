<?php
namespace Xls\Record;

class PrintHeaders extends AbstractRecord
{
    const NAME = 'PRINTHEADERS';
    const ID = 0x2A;

    /**
     * @param $printRowColHeaders
     *
     * @return string
     */
    public function getData($printRowColHeaders)
    {
        $data = pack("v", intval($printRowColHeaders));

        return $this->getFullRecord($data);
    }
}
