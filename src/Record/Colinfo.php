<?php

namespace Xls\Record;

class Colinfo extends AbstractRecord
{
    const NAME = 'COLINFO';
    const ID = 0x7D;

    /**
     * Generate the COLINFO biff record to define column widths
     *
     * @param array $colInfo This is the only parameter received and is composed of the following:
     *                0 => First formatted column,
     *                1 => Last formatted column,
     *                2 => Col width (8.43 is Excel default),
     *                3 => The optional XF format of the column,
     *                4 => Option flags.
     *                5 => Optional outline level
     *
     * @return string
     */
    public function getData($colInfo)
    {
        $colFirst = (isset($colInfo['col'])) ? $colInfo['col'] : 0;
        $colLast = (isset($colInfo['col2'])) ? $colInfo['col2'] : 0;

        $format = $colInfo['format'];
        $grbit = $colInfo['hidden'];
        $level = $colInfo['level'];

        $coldx = $colInfo['width'];
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
