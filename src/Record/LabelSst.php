<?php
namespace Xls\Record;

class LabelSst extends AbstractRecord
{
    const NAME = 'LABELSST';
    const ID = 0x00FD;

    /**
     * @param integer $row
     * @param integer $col
     * @param integer $strIdx
     * @param null $format
     *
     * @return string
     */
    public function getData($row, $col, $strIdx, $format = null)
    {
        $xf = $this->xf($format);

        $data = pack('vvvV', $row, $col, $xf, $strIdx);

        return $this->getFullRecord($data);
    }
}
