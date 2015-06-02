<?php
namespace Xls\Record;

use Xls\BIFFwriter;
use Xls\Format as XlsFormat;

abstract class AbstractRecord
{
    const ID = 0x00;

    /**
     * @var int
     */
    protected $byteOrder;

    /**
     * AbstractRecord constructor.
     *
     * @param int $byteOrder
     */
    public function __construct(
        $byteOrder = BIFFwriter::BYTE_ORDER_LE
    ) {
        $this->byteOrder = $byteOrder;
    }

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
        return (is_object($format)) ? $format->getXfIndex(): 0x0F;
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
}
