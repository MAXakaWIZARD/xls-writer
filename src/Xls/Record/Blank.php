<?php
namespace Xls\Record;

class Blank extends AbstractRecord
{
    const NAME = 'BLANK';
    const ID = 0x0201;
    const LENGTH = 0x06;

    /**
     * @param $row
     * @param $col
     * @param $format
     *
     * @return string
     */
    public function getData($row, $col, $format)
    {
        $xf = $this->xf($format);

        $data = pack("vvv", $row, $col, $xf);

        return $this->getHeader() . $data;
    }
}
