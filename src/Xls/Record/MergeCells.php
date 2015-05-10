<?php

namespace Xls\Record;

class MergeCells extends AbstractRecord
{
    const NAME = 'MERGECELLS';
    const ID = 0x00E5;
    const LENGTH = 0x02;

    /**
     * Generate the MERGECELLS biff record
     *
     * @param array $ranges
     *
     * @return string
     */
    public function getData($ranges)
    {
        $rangesCount = count($ranges);

        $data = pack('v', $rangesCount);
        foreach ($ranges as $range) {
            $data .= pack('vvvv', $range[0], $range[2], $range[1], $range[3]);
        }

        $extraLength = $rangesCount * 8;

        return $this->getHeader($extraLength) . $data;
    }
}
