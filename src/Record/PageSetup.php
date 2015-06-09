<?php

namespace Xls\Record;

use Xls\Worksheet;

class PageSetup extends AbstractRecord
{
    const NAME = 'PAGESETUP';
    const ID = 0xA1;

    /**
     * @param Worksheet $sheet
     *
     * @return string
     */
    public function getData($sheet)
    {
        $printSetup = $sheet->getPrintSetup();

        $iPaperSize = $printSetup->getPaperSize(); // Paper size
        $iScale = $printSetup->getPrintScale(); // Print scaling factor
        $iPageStart = 0x01; // Starting page number
        $iFitWidth = $printSetup->getFitWidth(); // Fit to number of pages wide
        $iFitHeight = $printSetup->getFitHeight(); // Fit to number of pages high
        $iRes = 0x0258; // Print resolution
        $iVRes = 0x0258; // Vertical print resolution

        $margin = $printSetup->getMargin();
        $numHdr = $margin->getHead(); // Header Margin
        $numFtr = $margin->getFoot(); // Footer Margin

        $numHdr = pack("d", $numHdr);
        $numFtr = pack("d", $numFtr);

        $data = pack(
            "vvvvvvvv",
            $iPaperSize,
            $iScale,
            $iPageStart,
            $iFitWidth,
            $iFitHeight,
            $this->calcGrbit($sheet),
            $iRes,
            $iVRes
        );
        $data .= $numHdr . $numFtr;

        $iCopies = 0x01; // Number of copies
        $data .= pack("v", $iCopies);

        return $this->getFullRecord($data);
    }

    /**
     * @param Worksheet $worksheet
     *
     * @return int
     */
    protected function calcGrbit(Worksheet $worksheet)
    {
        $fLeftToRight = 0x0; // Print over then down
        $fLandscape = $worksheet->getPrintSetup()->getOrientation(); // Page orientation
        $fNoPls = 0x0; // Setup not read from printer
        $fNoColor = 0x0; // Print black and white
        $fDraft = 0x0; // Print draft quality
        $fNotes = 0x0; // Print notes
        $fNoOrient = 0x0; // Orientation not set
        $fUsePage = 0x0; // Use custom starting page

        $grbit = $fLeftToRight;
        $grbit |= $fLandscape << 1;
        $grbit |= $fNoPls << 2;
        $grbit |= $fNoColor << 3;
        $grbit |= $fDraft << 4;
        $grbit |= $fNotes << 5;
        $grbit |= $fNoOrient << 6;
        $grbit |= $fUsePage << 7;

        return $grbit;
    }
}
