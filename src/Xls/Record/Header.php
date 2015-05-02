<?php

namespace Xls\Record;

use Xls\Biff8;

class Header extends AbstractRecord
{
    const NAME = 'HEADER';
    const ID = 0x0014;
    const LENGTH = 0x00;

    /**
     * Generate HEADER record
     *
     * @param $text
     *
     * @return string
     */
    public function getData($text)
    {
        $cch = strlen($text);
        if ($this->version === Biff8::VERSION) {
            $length = 3 + $cch;
            $encoding = 0x0;
            $data = pack("vC", $cch, $encoding);
        } else {
            $length = 1 + $cch;
            $data = pack("C", $cch);
        }

        return $this->getHeader($length) . $data . $text;
    }
}
