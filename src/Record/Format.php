<?php
namespace Xls\Record;

class Format extends AbstractRecord
{
    const NAME = 'FORMAT';
    const ID = 0x041E;

    /**
     * Generate FORMAT record for non "built-in" numerical formats.
     *
     * @param string $format Custom format string
     * @param integer $formatIndex   Format index code
     * @return string
     */
    public function getData($format, $formatIndex)
    {
        $formatLen = strlen($format);
        $length = 5 + $formatLen;

        if (function_exists('iconv')
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

        $data = pack("vvC", $formatIndex, $cch, $encoding);

        return $this->getHeader($length) . $data . $format;
    }
}
