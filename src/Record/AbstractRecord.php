<?php
namespace Xls\Record;

use Xls\Format as XlsFormat;

abstract class AbstractRecord
{
    const ID = 0x00;

    /**
     * Returns record header data ready for writing
     * @param int $length
     * @return string
     */
    public static function getHeader($length = 0)
    {
        return pack("vv", static::ID, $length);
    }

    /**
     * @param XlsFormat|null $format
     *
     * @return int
     */
    protected function xf($format)
    {
        return (is_object($format)) ? $format->getXfIndex() : 0x0F;
    }

    /**
     * returns full record data: header + data
     * @param $data
     *
     * @return string
     */
    protected function getFullRecord($data = '')
    {
        return $this->getHeader(strlen($data)) . $data;
    }

    /**
     * @param string $type
     * @param array $params
     *
     * @return mixed
     */
    protected function getSubRecord($type, array $params = array())
    {
        $callable = array("\\Xls\\Subrecord\\$type", 'getData');

        return call_user_func_array($callable, $params);
    }
}
