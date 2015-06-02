<?php
namespace Xls\Record;

class Header extends AbstractRecord
{
    const NAME = 'HEADER';
    const ID = 0x0014;

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
        $encoding = 0x0;
        $data = pack("vC", $cch, $encoding);
        $data .= $text;

        return $this->getFullRecord($data);
    }
}
