<?php

namespace Xls\Record;

class Obj extends AbstractRecord
{
    const NAME = 'OBJ';
    const ID = 0x005d;
    const LENGTH = 0x3c;

    /**
     * Generate the OBJ record that precedes an IMDATA record. This could be generalise
     * to support other Excel objects.
     *
     * @param integer $colL Column containing upper left corner of object
     * @param integer $dxL  Distance from left side of cell
     * @param integer $rwT  Row containing top left corner of object
     * @param integer $dyT  Distance from top of cell
     * @param integer $colR Column containing lower right corner of object
     * @param integer $dxR  Distance from right of cell
     * @param integer $rwB  Row containing bottom right corner of object
     * @param integer $dyB  Distance from bottom of cell
     * @return string
     */
    public function getData($colL, $dxL, $rwT, $dyT, $colR, $dxR, $rwB, $dyB)
    {
        $cObj = 0x0001; // Count of objects in file (set to 1)
        $OT = 0x0008; // Object type. 8 = Picture
        $id = 0x0001; // Object ID
        $grbit = 0x0614; // Option flags

        $cbMacro = 0x0000; // Length of FMLA structure
        $Reserved1 = 0x0000; // Reserved
        $Reserved2 = 0x0000; // Reserved

        $icvBack = 0x09; // Background colour
        $icvFore = 0x09; // Foreground colour
        $fls = 0x00; // Fill pattern
        $fAuto = 0x00; // Automatic fill
        $icv = 0x08; // Line colour
        $lns = 0xff; // Line style
        $lnw = 0x01; // Line weight
        $fAutoB = 0x00; // Automatic border
        $frs = 0x0000; // Frame style
        $cf = 0x0009; // Image format, 9 = bitmap
        $Reserved3 = 0x0000; // Reserved
        $cbPictFmla = 0x0000; // Length of FMLA structure
        $Reserved4 = 0x0000; // Reserved
        $grbit2 = 0x0001; // Option flags
        $Reserved5 = 0x0000; // Reserved

        $data = pack("V", $cObj);
        $data .= pack("v", $OT);
        $data .= pack("v", $id);
        $data .= pack("v", $grbit);
        $data .= pack("v", $colL);
        $data .= pack("v", $dxL);
        $data .= pack("v", $rwT);
        $data .= pack("v", $dyT);
        $data .= pack("v", $colR);
        $data .= pack("v", $dxR);
        $data .= pack("v", $rwB);
        $data .= pack("v", $dyB);
        $data .= pack("v", $cbMacro);
        $data .= pack("V", $Reserved1);
        $data .= pack("v", $Reserved2);
        $data .= pack("C", $icvBack);
        $data .= pack("C", $icvFore);
        $data .= pack("C", $fls);
        $data .= pack("C", $fAuto);
        $data .= pack("C", $icv);
        $data .= pack("C", $lns);
        $data .= pack("C", $lnw);
        $data .= pack("C", $fAutoB);
        $data .= pack("v", $frs);
        $data .= pack("V", $cf);
        $data .= pack("v", $Reserved3);
        $data .= pack("v", $cbPictFmla);
        $data .= pack("v", $Reserved4);
        $data .= pack("v", $grbit2);
        $data .= pack("V", $Reserved5);

        return $this->getHeader() . $data;
    }
}
