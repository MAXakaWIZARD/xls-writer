<?php

namespace Xls\Record;

class Externsheet extends AbstractRecord
{
    const NAME = 'EXTERNSHEET';
    const ID = 0x0017;

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
        $data .= $sheetName;

        return $this->getFullRecord($data);
    }

    /**
     * @param $sheetName
     * @param $currentSheetName
     *
     * @return string
     */
    public function getDataForCurrentSheet($sheetName, $currentSheetName)
    {
        if ($currentSheetName != $sheetName) {
            return $this->getData($sheetName);
        }

        $cch = 1; // The following byte
        $rgch = 0x02; // Self reference

        $data = pack("CC", $cch, $rgch);

        return $this->getFullRecord($data);
    }

    /**
     * @param $refs
     *
     * @return string
     */
    public function getDataForReferences($refs)
    {
        $refCount = count($refs);
        $data = pack('v', $refCount);

        foreach ($refs as $ref) {
            $data .= $ref;
        }

        return $this->getFullRecord($data);
    }
}
