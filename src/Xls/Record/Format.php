<?php

namespace Xls\Record;

use Xls\Biff8;

class Format extends AbstractRecord
{
    const NAME = 'FORMAT';
    const ID = 0x041E;
    const LENGTH = 0x00;

    /**
     * Generate FORMAT record for non "built-in" numerical formats.
     *
     * @param int $version BIFF version
     * @param string $format Custom format string
     * @param integer $formatIndex   Format index code
     * @return string
     */
    public function getData($version, $format, $formatIndex)
    {
        $formatLen = strlen($format);
        if ($version === Biff8::VERSION) {
            $length = 5 + $formatLen;
        } else {
            $length = 3 + $formatLen;
        }

        if ($version === Biff8::VERSION
            && function_exists('iconv')
        ) {
            // Encode format String
            if (mb_detect_encoding($format, 'auto') !== 'UTF-16LE') {
                $format = iconv(mb_detect_encoding($format, 'auto'), 'UTF-16LE', $format);
            }
            $encoding = 1;
            $cch = function_exists('mb_strlen') ? mb_strlen($format, 'UTF-16LE') : ($formatLen / 2);
        } else {
            $encoding = 0;
            $cch = $formatLen;
        }

        if ($version === Biff8::VERSION) {
            $data = pack("vvC", $formatIndex, $cch, $encoding);
        } else {
            $data = pack("vC", $formatIndex, $cch);
        }

        return $this->getHeader($length) . $data . $format;
    }
}
