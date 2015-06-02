<?php

namespace Xls\Record;

class Window1 extends AbstractRecord
{
    const NAME = 'WINDOW1';
    const ID = 0x003D;

    /**
     * @param $selectedSheetsCount Number of workbook tabs selected
     * @param $firstSheet 1st displayed worksheet
     * @param $activeSheet 1st displayed worksheet
     *
     * @return string
     */
    public function getData($selectedSheetsCount, $firstSheet, $activeSheet)
    {
        $xWn = 0x0000; // Horizontal position of window
        $yWn = 0x0000; // Vertical position of window
        $dxWn = 0x25BC; // Width of window
        $dyWn = 0x1572; // Height of window

        $grbit = 0x0038; // Option flags
        $wTabRatio = 0x0258; // Tab to scrollbar ratio

        $data = pack(
            "vvvvvvvvv",
            $xWn,
            $yWn,
            $dxWn,
            $dyWn,
            $grbit,
            $activeSheet,
            $firstSheet,
            $selectedSheetsCount,
            $wTabRatio
        );

        return $this->getFullRecord($data);
    }
}
