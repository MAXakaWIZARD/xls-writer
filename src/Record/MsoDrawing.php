<?php

namespace Xls\Record;

class MsoDrawing extends AbstractRecord
{
    const NAME = 'MSODRAWING';
    const ID = 0x00EC;

    public function getData($hexStrData)
    {
        $data = pack('H*', str_replace(' ', '', $hexStrData));

        return $this->getFullRecord($data);
    }
}
