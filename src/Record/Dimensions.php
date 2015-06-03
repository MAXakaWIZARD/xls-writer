<?php
namespace Xls\Record;

use Xls\Range;

class Dimensions extends AbstractRecord
{
    const NAME = 'DIMENSIONS';
    const ID = 0x0200;

    /**
     * @param Range $range
     *
     * @return string
     */
    public function getData(Range $range)
    {
        $reserved = 0x00;

        $data = pack("VV", $range->getRowFrom(), $range->getRowTo() + 1);
        $data .= pack(
            "vvv",
            $range->getColFrom(),
            $range->getColTo() + 1,
            $reserved
        );

        return $this->getFullRecord($data);
    }
}
