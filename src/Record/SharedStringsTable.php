<?php

namespace Xls\Record;

use Xls\SharedStringsTable as SST;

class SharedStringsTable extends AbstractRecord
{
    const NAME = 'SST';
    const ID = 0x00fc;

    /**
     * Write all of the workbooks strings into an indexed array.
     *
     * The Excel documentation says that the SST record should be followed by an
     * EXTSST record. The EXTSST record is a hash table that is used to optimise
     * access to SST. However, despite the documentation it doesn't seem to be
     * required so we will ignore it.
     *
     * @param SST $sst
     * @return string
     */
    public function getData(SST $sst)
    {
        $data = pack("VV", $sst->getTotalCount(), $sst->getUniqueCount());

        $length = strlen($data);
        $blockSizes = $sst->getBlocksSizes();
        if (!empty($blockSizes)) {
            $length += array_shift($blockSizes);
        }

        return $this->getHeader($length) . $data;
    }
}
