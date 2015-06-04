<?php

namespace Xls\Record;

use Xls\Range;

class MergeCells extends AbstractRecord
{
    const NAME = 'MERGECELLS';
    const ID = 0x00E5;

    /**
     * Generate the MERGECELLS biff record
     *
     * @param Range[] $ranges
     *
     * @return string
     */
    public function getData($ranges)
    {
        $data = $this->getSubRecord('Range', array($ranges));

        return $this->getFullRecord($data);
    }
}
