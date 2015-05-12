<?php
namespace Xls\Record;

class NameLong extends NameShort
{
    const NAME = 'NAME';
    const ID = 0x0018;
    const LENGTH = 0x3D;

    // Length of text definition
    const CCE = 0x002e;

    const UNKNOWN_08 = 0x8008;

    /**
     * @param $index
     * @param $rowmin
     * @param $rowmax
     * @param $colmin
     * @param $colmax
     *
     * @return string
     */
    protected function getExtraData($index, $rowmin, $rowmax, $colmin, $colmax)
    {
        $unknown01 = 0x29;
        $unknown02 = 0x002b;
        $data = pack("Cv", $unknown01, $unknown02);

        $common = $this->getRowColDefCommonData($index);

        // Column definition
        $data .= $common;
        $data .= pack("vv", 0x0000, 0x3fff);
        $data .= pack("CC", $colmin, $colmax);

        // Row definition
        $data .= $common;
        $data .= pack("vv", $rowmin, $rowmax);
        $data .= pack("CC", 0x00, 0xff);

        // End of data
        $data .= pack("C", 0x10);

        return $data;
    }
}
