<?php
namespace Xls\Record;

class Window2 extends AbstractRecord
{
    const NAME = 'WINDOW2';
    const ID = 0x023E;

    /**
     * @param \Xls\Worksheet $worksheet
     *
     * @return string
     */
    public function getData($worksheet)
    {
        if ($worksheet->isBiff5()) {
            $length = 0x0A; // Number of bytes to follow
        } else {
            $length = 0x12;
        }

        $rwTop = 0x00; // Top row visible in window
        $colLeft = 0x00; // Leftmost column visible in window

        // The options flags that comprise $grbit
        $fDspFmla = 0; // 0 - bit
        $fDspGrid = intval($worksheet->areScreenGridLinesVisible()); // 1
        $fDspRwCol = 1; // 2
        $fFrozen = intval($worksheet->isFrozen()); // 3
        $fDspZeros = 1; // 4
        $fDefaultHdr = 1; // 5
        $fArabic = intval($worksheet->isArabic()); // 6
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

        $data = pack("vvv", $grbit, $rwTop, $colLeft);

        if ($worksheet->isBiff5()) {
            $rgbHdr = 0x00; // Row/column heading and gridline color
            $data .= pack("V", $rgbHdr);
        } else {
            $rgbHdr = 0x0040; // Row/column heading and gridline color index
            $zoomFactorPageBreak = 0x00;
            $zoomFactorNormal = 0x00;
            $data .= pack("vvvvV", $rgbHdr, 0x00, $zoomFactorPageBreak, $zoomFactorNormal, 0x00);
        }

        return $this->getHeader($length) . $data;
    }
}
