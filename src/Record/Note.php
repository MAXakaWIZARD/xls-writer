<?php
namespace Xls\Record;

use Xls\StringUtils;

class Note extends AbstractRecord
{
    const NAME = 'NOTE';
    const ID = 0x001C;

    /**
     * Generate a note associated with the cell given by the row and column.
     * NOTE records don't have a length limit
     * @param integer $row
     * @param integer $col
     * @param integer $objId
     *
     * @return string
     */
    public function getData($row, $col, $objId)
    {
        $grbit = 0x00;
        $data = pack("vvvv", $row, $col, $grbit, $objId);
        $author = 'xls-writer';
        $data .= StringUtils::toBiff8UnicodeLong($author);

        return $this->getFullRecord($data);
    }
}
