<?php

namespace Xls\Record;

use Xls\Range;
use Xls\Margin;

class ObjPicture extends Obj
{
    const TYPE = 0x08;

    /**
     * Generate the OBJ record that precedes an IMDATA record. This could be generalise
     * to support other Excel objects.
     *
     * @param integer $objectId
     * @param Range $area Picture position area
     * @param Margin $margin  Margins from cell sides
     * @return string
     */
    public function getData($objectId, Range $area, Margin $margin)
    {
        $objCount = 0x01; // Count of objects in file (set to 1)
        $grbit = 0x0614; // Option flags
        $data = pack("Vv3", $objCount, static::TYPE, $objectId, $grbit);

        $cbMacro = 0x00; // Length of FMLA structure
        $reserved = 0x00; // Reserved

        $icvBack = 0x09; // Background colour
        $icvFore = 0x09; // Foreground colour
        $fls = 0x00; // Fill pattern
        $fAuto = 0x00; // Automatic fill
        $icv = 0x08; // Line colour
        $lns = 0xff; // Line style
        $lnw = 0x01; // Line weight
        $fAutoB = 0x00; // Automatic border
        $frs = 0x00; // Frame style
        $imageFormat = 0x09; // Image format, 9 = bitmap
        $cbPictFmla = 0x00; // Length of FMLA structure
        $grbit2 = 0x01; // Option flags

        $data .= $this->packArea($area, $margin);

        $data .= pack("v", $cbMacro);
        $data .= pack("Vv", $reserved, $reserved);
        $data .= pack("C2", $icvBack, $icvFore);
        $data .= pack("C6", $fls, $fAuto, $icv, $lns, $lnw, $fAutoB);
        $data .= pack("vV", $frs, $imageFormat);
        $data .= pack("v4", $reserved, $cbPictFmla, $reserved, $grbit2);

        $data .= $this->getFtEndSubrecord();

        return $this->getFullRecord($data);
    }

    /**
     * @param Range  $area
     * @param Margin $margin
     *
     * @return string
     */
    protected function packArea(Range $area, Margin $margin)
    {
        $data = pack("v2", $area->getColFrom(), $margin->getLeft());
        $data .= pack("v2", $area->getRowFrom(), $margin->getTop());
        $data .= pack("v2", $area->getColTo(), $margin->getRight());
        $data .= pack("v2", $area->getRowTo(), $margin->getBottom());

        return $data;
    }
}
