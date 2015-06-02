<?php
namespace Xls\Record;

class Dimensions extends AbstractRecord
{
    const NAME = 'DIMENSIONS';
    const ID = 0x0200;

    /**
     * @param $rowMin
     * @param $rowMax
     * @param $colMin
     * @param $colMax
     *
     * @return string
     */
    public function getData($rowMin, $rowMax, $colMin, $colMax)
    {
        $reserved = 0x00;

        $data = pack("VV", $rowMin, $rowMax);
        $data .= pack(
            "vvv",
            $colMin,
            $colMax,
            $reserved
        );

        return $this->getFullRecord($data);
    }
}
