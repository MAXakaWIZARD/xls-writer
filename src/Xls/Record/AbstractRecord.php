<?php
namespace Xls\Record;

use Xls\BIFFwriter;
use Xls\Biff5;
use Xls\Biff8;
use Xls\Format as XlsFormat;

abstract class AbstractRecord
{
    const ID = 0x00;
    const LENGTH = 0x00;

    /**
     * @var int
     */
    protected $version;

    /**
     * @var int
     */
    protected $byteOrder;

    /**
     * AbstractRecord constructor.
     *
     * @param int $version BIFF version
     * @param int $byteOrder
     */
    public function __construct(
        $version,
        $byteOrder = BIFFwriter::BYTE_ORDER_LE
    ) {
        $this->version = $version;
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

        if (is_null($extraParam)) {
            return pack("vv", static::ID, $length);
        } else {
            return pack("vvv", static::ID, $length, $extraParam);
        }
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
     *
     */
    protected function isBiff5()
    {
        return $this->version === Biff5::VERSION;
    }

    /**
     *
     */
    protected function isBiff8()
    {
        return $this->version === Biff8::VERSION;
    }
}
