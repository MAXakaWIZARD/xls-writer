<?php
namespace Xls\Record;

use Xls\Worksheet;

class WsBool extends AbstractRecord
{
    const NAME = 'WSBOOL';
    const ID = 0x0081;
    const LENGTH = 0x02;

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
        if ($sheet->outlineStyle) {
            $grbit |= 0x0020; // Auto outline styles
        }
        if ($sheet->outlineBelow) {
            $grbit |= 0x0040; // Outline summary below
        }
        if ($sheet->outlineRight) {
            $grbit |= 0x0080; // Outline summary right
        }
        if ($sheet->fitPage) {
            $grbit |= 0x0100; // Page setup fit to page
        }
        if ($sheet->isOutlineOn()) {
            $grbit |= 0x0400; // Outline symbols displayed
        }

        $data = pack("v", $grbit);

        return $this->getHeader() . $data;
    }
}