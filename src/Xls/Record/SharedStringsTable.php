<?php

namespace Xls\Record;

class SharedStringsTable extends AbstractRecord
{
    const NAME = 'SST';
    const ID = 0x00fc;
    const LENGTH = 0x08;

    /**
     * Write all of the workbooks strings into an indexed array.
     *
     * The Excel documentation says that the SST record should be followed by an
     * EXTSST record. The EXTSST record is a hash table that is used to optimise
     * access to SST. However, despite the documentation it doesn't seem to be
     * required so we will ignore it.
     *
     * @param $blockSizes
     * @param $strTotal
     * @param $strUnique
     * @return string
     */
    public function getData($blockSizes, $strTotal, $strUnique)
    {
        // The SST record is required even if it contains no strings. Thus we will
        // always have a length
        //
        if (!empty($blockSizes)) {
            $extraLength = array_shift($blockSizes);
        } else {
            // No strings
            $extraLength = 0;
        }

        // Write the SST block header information
        $data = pack("VV", $strTotal, $strUnique);

        return $this->getHeader($extraLength) . $data;
    }
}
