<?php

namespace Xls\Record;

use Xls\Range;

class ObjPicture extends Obj
{
    const TYPE = 0x08;

    /**
     * Generate the OBJ record that precedes an IMDATA record. This could be generalise
     * to support other Excel objects.
     *
     * @param integer $objectId
     * @param Range $area Picture position area
     * @param integer $dxL  Distance from left side of cell
     * @param integer $dyT  Distance from top of cell
     * @param integer $dxR  Distance from right of cell
     * @param integer $dyB  Distance from bottom of cell
     * @return string
     */
    public function getData($objectId, $area, $dxL, $dyT, $dxR, $dyB)
    {
        $type = static::TYPE;
        $cObj = 0x0001; // Count of objects in file (set to 1)
        $id = $objectId; // Object ID
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
        $reserved3 = 0x0000; // Reserved
        $cbPictFmla = 0x0000; // Length of FMLA structure
        $reserved4 = 0x0000; // Reserved
        $grbit2 = 0x0001; // Option flags

        $data = pack("V", $cObj);
        $data .= pack("v", $type);
        $data .= pack("v", $id);
        $data .= pack("v", $grbit);

        $data .= pack("v", $area->getColFrom());
        $data .= pack("v", $dxL);
        $data .= pack("v", $area->getRowFrom());
        $data .= pack("v", $dyT);
        $data .= pack("v", $area->getColTo());
        $data .= pack("v", $dxR);
        $data .= pack("v", $area->getRowTo());
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
        $data .= pack("v", $reserved3);
        $data .= pack("v", $cbPictFmla);
        $data .= pack("v", $reserved4);
        $data .= pack("v", $grbit2);

        $data .= $this->getFtEndSubrecord();

        return $this->getFullRecord($data);
    }
}
