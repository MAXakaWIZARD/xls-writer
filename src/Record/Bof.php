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
        $length = 0x10;
        $unknown = pack("VV", 0x000100D1, 0x00000406);
        $build = 0x0DBB;
        $year = 0x07CC;

        $data = pack("vvvv", $this->version, $type, $build, $year);

        return $this->getHeader($length) . $data . $unknown;
    }
}
