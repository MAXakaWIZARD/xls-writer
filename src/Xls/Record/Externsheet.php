<?php

namespace Xls\Record;

class Externsheet extends AbstractRecord
{
    const NAME = 'EXTERNSHEET';
    const ID = 0x0017;
    const LENGTH = 0x02;

    /**
     * Writes the Excel BIFF EXTERNSHEET record. These references are used by
     * formulas. NAME record is required to define the print area and the repeat
     * rows and columns.
     *
     * @param string $sheetName Worksheet name
     * @return string
     */
    public function getData($sheetName)
    {
        $cch = strlen($sheetName);
        $rgch = 0x03; // Filename encoding

        $data = pack("CC", $cch, $rgch);

        return $this->getHeader($cch) . $data . $sheetName;
    }

    /**
     * @param $refs
     *
     * @return string
     */
    public function getDataForReferences($refs)
    {
        $refCount = count($refs);
        $extraLength = 6 * $refCount;

        $data = pack('v', $refCount);
        foreach ($refs as $ref) {
            $data .= $ref;
        }

        return $this->getHeader($extraLength) . $data;
    }
}
