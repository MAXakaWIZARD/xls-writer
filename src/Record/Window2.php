<?php
namespace Xls\Record;

use Xls\Worksheet;

class Window2 extends AbstractRecord
{
    const NAME = 'WINDOW2';
    const ID = 0x023E;

    /**
     * @param Worksheet $worksheet
     *
     * @return string
     */
    public function getData(Worksheet $worksheet)
    {
        $rwTop = 0x00; // Top row visible in window
        $colLeft = 0x00; // Leftmost column visible in window

        $data = pack("vvv", $this->calcGrbit($worksheet), $rwTop, $colLeft);

        $rgbHdr = 0x0040; // Row/column heading and gridline color index
        $zoomFactorPageBreak = 0x00;
        $zoomFactorNormal = 0x00;
        $data .= pack("vvvvV", $rgbHdr, 0x00, $zoomFactorPageBreak, $zoomFactorNormal, 0x00);

        return $this->getFullRecord($data);
    }

    /**
     * @param Worksheet $worksheet
     *
     * @return int
     */
    protected function calcGrbit(Worksheet $worksheet)
    {
        $fDspFmla = 0; // 0 - bit
        $fDspGrid = intval($worksheet->areGridLinesVisible()); // 1
        $fDspRwCol = 1; // 2
        $fFrozen = intval($worksheet->isFrozen()); // 3
        $fDspZeros = 1; // 4
        $fDefaultHdr = 1; // 5
        $fArabic = intval($worksheet->isRtl()); // 6
        $fDspGuts = intval($worksheet->isOutlineOn()); // 7
        $fFrozenNoSplit = 0; // 0 - bit
        $fSelected = intval($worksheet->isSelected()); // 1
        $fPaged = 1; // 2

        $grbit = $fDspFmla;
        $grbit |= $fDspGrid << 1;
        $grbit |= $fDspRwCol << 2;
        $grbit |= $fFrozen << 3;
        $grbit |= $fDspZeros << 4;
        $grbit |= $fDefaultHdr << 5;
        $grbit |= $fArabic << 6;
        $grbit |= $fDspGuts << 7;
        $grbit |= $fFrozenNoSplit << 8;
        $grbit |= $fSelected << 9;
        $grbit |= $fPaged << 10;

        return $grbit;
    }
}
