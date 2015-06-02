<?php
namespace Xls\Record;

use Xls\BIFFwriter;
use Xls\Biff8;
use Xls\Format as XlsFormat;

abstract class AbstractRecord
{
    const ID = 0x00;
    const LENGTH = 0x00;
    const HEADER_SIZE = 4;

    /**
     * @var int
     */
    protected $version = Biff8::VERSION;

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
     * @param int $extraLength
     * @param int $extraParam
     * @return string
     */
    public static function getHeader($extraLength = 0, $extraParam = null)
    {
        $length = static::LENGTH + $extraLength;

        $header = pack("vv", static::ID, $length);

        if (!is_null($extraParam)) {
            $header .= pack("v", $extraParam);
        }

        return $header;
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
    protected function getFullRecord($data)
    {
        return $this->getHeader(strlen($data)) . $data;
    }
}
