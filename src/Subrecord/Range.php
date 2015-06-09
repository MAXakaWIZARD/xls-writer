<?php

namespace Xls\Subrecord;

use Xls\Range as CellRange;

class Range
{
    /**
     * @param CellRange[] $ranges
     * @param bool $includeCount
     * @return string
     */
    public static function getData($ranges, $includeCount = true)
    {
        $data = '';

        if ($includeCount) {
            $rangesCount = count($ranges);
            $data .= pack('v', $rangesCount);
        }

        foreach ($ranges as $range) {
            $data .= pack(
                'v4',
                $range->getRowFrom(),
                $range->getRowTo(),
                $range->getColFrom(),
                $range->getColTo()
            );
        }

        return $data;
    }
}
