<?php
namespace Xls\Record;

use Xls\BIFFwriter;

class Number extends AbstractRecord
{
    const NAME = 'NUMBER';
    const ID = 0x0203;
    const LENGTH = 0x0E;

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

        $xlDouble = pack("d", $num);
        if ($this->byteOrder === BIFFwriter::BYTE_ORDER_BE) {
            $xlDouble = strrev($xlDouble);
        }

        return $this->getHeader() . $data . $xlDouble;
    }
}
