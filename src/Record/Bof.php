<?php
namespace Xls\Record;

use Xls\Biff8;

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
        $build = 0x0DBB;
        $year = 0x07CC;

        $data = pack("vvvv", Biff8::VERSION, $type, $build, $year);
        $unknown = pack("VV", 0x000100D1, 0x00000406);
        $data .= $unknown;

        return $this->getFullRecord($data);
    }
}
