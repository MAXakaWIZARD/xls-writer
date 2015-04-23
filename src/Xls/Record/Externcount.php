<?php

namespace Xls\Record;

class Externcount extends AbstractRecord
{
    const NAME = 'EXTERNCOUNT';
    const ID = 0x0016;
    const LENGTH = 0x02;

    /**
     * Write BIFF record EXTERNCOUNT to indicate the number of external sheet
     * references in the workbook.
     *
     * Excel only stores references to external sheets that are used in NAME.
     * The workbook NAME record is required to define the print area and the repeat
     * rows and columns.
     *
     * @param integer $externalRefsCount Number of external references
     * @return string
     */
    public function getData($externalRefsCount)
    {
        $data = pack("v", $externalRefsCount);

        return $this->getHeader() . $data;
    }
}
