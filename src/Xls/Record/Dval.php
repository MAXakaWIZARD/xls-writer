<?php

namespace Xls\Record;

class Dval extends AbstractRecord
{
    const NAME = 'DVAL';
    const ID = 0x01b2;
    const LENGTH = 0x12;

    /**
     * Generate the DVAL biff record
     * @param $dv
     *
     * @return string
     */
    public function getData($dv)
    {
        $grbit = 0x0002; // Prompt box at cell, no cached validity data at DV records
        $horPos = 0x00000000; // Horizontal position of prompt box, if fixed position
        $verPos = 0x00000000; // Vertical position of prompt box, if fixed position
        $objId = 0xffffffff; // Object identifier of drop down arrow object, or -1 if not visible

        $data = pack(
            'vVVVV',
            $grbit,
            $horPos,
            $verPos,
            $objId,
            count($dv)
        );

        return $this->getHeader() . $data;
    }
}
