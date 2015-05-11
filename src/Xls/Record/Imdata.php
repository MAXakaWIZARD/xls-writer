<?php
namespace Xls\Record;

class Imdata extends AbstractRecord
{
    const NAME = 'IMDATA';
    const ID = 0x007f;

    /**
     * Generate IMDATA record
     * @param $size
     * @param $data
     *
     * @return string
     */
    public function getData($size, $data)
    {
        $length = 8 + $size;
        $cf = 0x09;
        $env = 0x01;

        $header = pack("vvvvV", static::ID, $length, $cf, $env, $size);

        return $header . $data;
    }
}
