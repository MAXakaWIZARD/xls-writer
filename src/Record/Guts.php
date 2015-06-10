<?php

namespace Xls\Record;

class Guts extends AbstractRecord
{
    const NAME = 'GUTS';
    const ID = 0x0080;

    /**
     * Generate the GUTS BIFF record. This is used to configure the gutter margins
     * where Excel outline symbols are displayed. The visibility of the gutters is
     * controlled by a flag in WSBOOL.
     *
     * @param $colInfo
     * @param $outlineRowLevel
     *
     * @return string
     */
    public function getData($colInfo, $outlineRowLevel)
    {
        $dxRwGut = 0x0000; // Size of row gutter
        $dxColGut = 0x0000; // Size of col gutter

        $rowLevel = $outlineRowLevel;
        $colLevel = 0;

        // Calculate the maximum column outline level. The equivalent calculation
        // for the row outline level is carried out in setRow().
        foreach ($colInfo as $col) {
            $colLevel = max($col['level'], $colLevel);
        }

        // Set the limits for the outline levels (0 <= x <= 7).
        $colLevel = max(0, min($colLevel, 7));

        // The displayed level is one greater than the max outline levels
        if ($rowLevel) {
            $rowLevel++;
        }
        if ($colLevel) {
            $colLevel++;
        }

        $data = pack("vvvv", $dxRwGut, $dxColGut, $rowLevel, $colLevel);

        return $this->getFullRecord($data);
    }
}
