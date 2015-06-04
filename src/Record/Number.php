<?php
namespace Xls\Record;

use Xls\BIFFwriter;

class Number extends AbstractRecord
{
    const NAME = 'NUMBER';
    const ID = 0x0203;

    /**
     * @param integer $row
     * @param integer $col
     * @param float $num
     * @param null $format
     *
     * @return string
     */
    public function getData($row, $col, $num, $format = null)
    {
        $xf = $this->xf($format);
        $data = pack("vvv", $row, $col, $xf);
        $data .= pack("d", $num);

        return $this->getFullRecord($data);
    }
}
