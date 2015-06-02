<?php
namespace Xls\Record;

class Eof extends AbstractRecord
{
    const NAME = 'EOF';
    const ID = 0x000A;

    /**
     * Generate EOF record to indicate the end of a BIFF stream
     * @return string
     */
    public function getData()
    {
        return $this->getFullRecord();
    }
}
