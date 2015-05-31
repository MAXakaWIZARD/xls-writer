<?php

namespace Xls\Record;

use Xls\Utils;

abstract class Obj extends AbstractRecord
{
    const NAME = 'OBJ';
    const ID = 0x005d;

    protected function getFtCmoSubrecord($objectId)
    {
        $grbit = 0x4011;
        $data = pack("vvv", static::TYPE, $objectId, $grbit);

        //reserved
        $data .= pack("VVV", 0, 0, 0);

        $header = pack("vv", 0x15, strlen($data));

        return $header . $data;
    }

    protected function getFtNtsSubrecord($guid = null)
    {
        $length = 0x16;
        $header = pack("vv", 0x0D, $length);

        $guid = (is_null($guid)) ? Utils::generateGuid() : $guid;
        $data = pack('H*', $guid);

        $fSharedNote = 0; //not shared
        $data .= pack('v', $fSharedNote);

        //reserved
        $data .= pack('vv', 0x10, 0);

        return $header . $data;
    }

    protected function getFtEndSubrecord()
    {
        return pack("vv", 0x00, 0x00);
    }
}
