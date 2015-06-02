<?php
namespace Xls\Record;

use Xls\StringUtils;

class Txo extends AbstractRecord
{
    const NAME = 'TXO';
    const ID = 0x01B6;

    /**
     * @param $text
     *
     * @return string
     */
    public function getData($text)
    {
        $charCount = StringUtils::countCharacters($text);
        $text = StringUtils::toBiff8UnicodeLongWoLenInfo($text);

        $grbit = 0x0212;
        $rotation = 0;
        $data = pack('vv', $grbit, $rotation);
        $data .= pack('vvv', 0, 0, 0); //reserved
        $txoRunsLength = 0x10;
        $data .= pack('vv', $charCount, $txoRunsLength);
        $data .= pack('V', 0); //reserved

        $result = $this->getFullRecord($data);

        $continue = new ContinueRecord();
        $result .= $continue->getDataRaw($text);

        $txoRunsData = pack('H*', '00000500');
        $txoRunsData .= pack('H*', '2F00');
        $txoRunsData .= pack('H*', '0C00');

        $lastRun = pack('v', $charCount);
        $lastRun .= pack('H*', '000000000200');

        $txoRunsData .= $lastRun;
        $result .= $continue->getDataRaw($txoRunsData);

        return $result;
    }
}
