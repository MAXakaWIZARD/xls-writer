<?php
namespace Xls\Record;

class Label extends AbstractRecord
{
    const NAME = 'LABEL';
    const ID = 0x0204;
    const LENGTH = 0x08;

    /**
     * @param      $row
     * @param      $col
     * @param      $str
     * @param null $format
     *
     * @return string
     */
    public function getData($row, $col, $str, $format = null)
    {
        $xf = $this->xf($format); // The cell format

        $strlen = strlen($str);

        $data = pack("vvvv", $row, $col, $xf, $strlen);

        return $this->getHeader($strlen) . $data . $str;
    }
}
