<?php
namespace Xls\Record;

class Boundsheet extends AbstractRecord
{
    const NAME = 'BOUNDSHEET';
    const ID = 0x0085;

    /**
     * Generate BOUNDSHEET record.
     *
     * @param string $sheetName Worksheet name
     * @param integer $offset    Location of worksheet BOF
     * @return string
     */
    public function getData($sheetName, $offset)
    {
        $grbit = 0x0000;

        if ($this->isBiff8()) {
            $length = 0x08 + strlen($sheetName);
            $cch = mb_strlen($sheetName, 'UTF-16LE');
            $data = pack("VvCC", $offset, $grbit, $cch, 0x1);
        } else {
            $length = 0x07 + strlen($sheetName);
            $cch = strlen($sheetName);
            $data = pack("VvC", $offset, $grbit, $cch);
        }

        return $this->getHeader($length) . $data . $sheetName;
    }
}
