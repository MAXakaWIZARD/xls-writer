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
        $rangesCount = count($ranges);

        $data = pack('v', $rangesCount);
        foreach ($ranges as $range) {
            $data .= pack(
                'vvvv',
                $range->getRowFrom(),
                $range->getRowTo(),
                $range->getColFrom(),
                $range->getColTo()
            );
        }

        return $this->getFullRecord($data);
    }
}
