<?php

namespace Xls\Record;

class Colinfo extends AbstractRecord
{
    const NAME = 'COLINFO';
    const ID = 0x7D;

    /**
     * Generate the COLINFO biff record to define column widths
     *
     * @param array $colArray This is the only parameter received and is composed of the following:
     *                0 => First formatted column,
     *                1 => Last formatted column,
     *                2 => Col width (8.43 is Excel default),
     *                3 => The optional XF format of the column,
     *                4 => Option flags.
     *                5 => Optional outline level
     *
     * @return string
     */
    public function getData($colArray)
    {
        $colFirst = (isset($colArray[0])) ? $colArray[0] : 0;
        $colLast = (isset($colArray[1])) ? $colArray[1] : 0;

        if (isset($colArray[2])) {
            $coldx = $colArray[2];
        } else {
            $coldx = 8.43;
        }

        if (isset($colArray[3])) {
            $format = $colArray[3];
        } else {
            $format = null;
        }

        if (isset($colArray[4])) {
            $grbit = $colArray[4];
        } else {
            $grbit = 0;
        }

        if (isset($colArray[5])) {
            $level = $colArray[5];
        } else {
            $level = 0;
        }

        $coldx += 0.72; // Fudge. Excel subtracts 0.72 !?
        $coldx *= 256; // Convert to units of 1/256 of a char

        $reserved = 0x00; // Reserved

        $level = max(0, min($level, 7));
        $grbit |= $level << 8;

        $data = pack(
            "vvvvvC",
            $colFirst,
            $colLast,
            $coldx,
            $this->xf($format),
            $grbit,
            $reserved
        );

        return $this->getFullRecord($data);
    }
}
