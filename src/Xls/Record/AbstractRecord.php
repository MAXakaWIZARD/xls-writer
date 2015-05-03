<?php

namespace Xls\Record;

use Xls\BIFFwriter;
use Xls\Biff5;

abstract class AbstractRecord
{
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
     * @param int $version
     * @param int $byteOrder
     */
    public function __construct(
        $version = Biff5::VERSION,
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
    public function getHeader($extraLength = 0, $extraParam = null)
    {
        $length = static::LENGTH + $extraLength;

        if (is_null($extraParam)) {
            return pack("vv", static::ID, $length);
        } else {
            return pack("vvv", static::ID, $length, $extraParam);
        }
    }
}
