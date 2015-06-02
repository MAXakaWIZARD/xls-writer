<?php
namespace Xls\Record;

class Imdata extends AbstractRecord
{
    const NAME = 'IMDATA';
    const ID = 0x007f;

    /**
     * Generate IMDATA record
     * @param $imgData
     *
     * @return string
     */
    public function getData($imgData)
    {
        $cf = 0x09;
        $env = 0x01;

        $size = strlen($imgData);
        $data = pack("vvV", $cf, $env, $size);
        $data .= $imgData;

        return $this->getFullRecord($data);
    }
}
