<?php
namespace Xls\Record;

use Xls\Worksheet;

class WsBool extends AbstractRecord
{
    const NAME = 'WSBOOL';
    const ID = 0x0081;

    /**
     * Generate the WSBOOL biff record
     * @param Worksheet $sheet
     *
     * @return string
     */
    public function getData(Worksheet $sheet)
    {
        $grbit = 0x0000;

        // Set the option flags
        $grbit |= 0x0001; // Auto page breaks visible
        if ($sheet->getOutlineStyle()) {
            $grbit |= 0x0020; // Auto outline styles
        }
        if ($sheet->getOutlineBelow()) {
            $grbit |= 0x0040; // Outline summary below
        }
        if ($sheet->getOutlineRight()) {
            $grbit |= 0x0080; // Outline summary right
        }
        if ($sheet->getPrintSetup()->isFitPage()) {
            $grbit |= 0x0100; // Page setup fit to page
        }
        if ($sheet->isOutlineOn()) {
            $grbit |= 0x0400; // Outline symbols displayed
        }

        $data = pack("v", $grbit);

        return $this->getFullRecord($data);
    }
}
