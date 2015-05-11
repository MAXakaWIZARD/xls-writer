<?php
namespace Xls\Record;

class Bof extends AbstractRecord
{
    const NAME = 'BOF';
    const ID = 0x0809;

    /**
     * Generate BOF record to indicate the beginning of a stream or
     * sub-stream in the BIFF file.
     *
     * @param integer $type Type of BIFF file to write: Workbook or Worksheet.
     * @return string
     */
    public function getData($type)
    {
        // According to the SDK $build and $year should be set to zero.
        // However, this throws a warning in Excel 5. So, use magic numbers.
        if ($this->isBiff5()) {
            $length = 0x08;
            $unknown = '';
            $build = 0x096C;
            $year = 0x07C9;
        } else {
            $length = 0x10;
            $unknown = pack("VV", 0x00000041, 0x00000006); //unknown last 8 bytes for BIFF8
            $build = 0x0DBB;
            $year = 0x07CC;
        }

        $data = pack("vvvv", $this->version, $type, $build, $year);

        return $this->getHeader($length) . $data . $unknown;
    }
}
