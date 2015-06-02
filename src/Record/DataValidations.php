<?php

namespace Xls\Record;

class DataValidations extends AbstractRecord
{
    const NAME = 'DATAVALIDATIONS';
    const ID = 0x01B2;

    /**
     * Generate the DVAL biff record
     * @param $dv
     *
     * @return string
     */
    public function getData($dv)
    {
        $grbit = 0x02; // Prompt box at cell, no cached validity data at DV records
        $horPos = 0x00; // Horizontal position of prompt box, if fixed position
        $verPos = 0x00; // Vertical position of prompt box, if fixed position
        $objId = 0xffffffff; // Object identifier of drop down arrow object, or -1 if not visible

        $data = pack(
            'vVVVV',
            $grbit,
            $horPos,
            $verPos,
            $objId,
            count($dv)
        );

        return $this->getFullRecord($data);
    }
}
