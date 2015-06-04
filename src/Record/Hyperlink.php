<?php
namespace Xls\Record;

use Xls\StringUtils;
use Xls\Range;

class Hyperlink extends AbstractRecord
{
    const NAME = 'HYPERLINK';
    const ID = 0x01B8;
    const STDLINK_GUID = "D0C9EA79F9BACE118C8200AA004BA90B";
    const MONIKER_GUID = "E0C9EA79F9BACE118C8200AA004BA90B";

    /**
     * @param Range $range
     * @param $url
     *
     * @return string
     */
    public function getData(Range $range, $url)
    {
        $url = StringUtils::toNullTerminatedWchar($url);

        $options = $this->getOptions($url);
        $data = $this->getCommonData($range, $options);
        $data .= pack("H*", static::MONIKER_GUID);
        $data .= pack("V", strlen($url));
        $data .= $url;

        return $this->getFullRecord($data);
    }

    protected function getOptions($url)
    {
        $options = 0x00;
        $options |= 1 << 0; //File link or URL
        $options |= 1 << 1; //Absolute path or URL

        return $options;
    }

    protected function getCommonData(Range $range, $options)
    {
        $data = $this->getSubRecord('Range', array(array($range), false));
        $data .= pack("H*", static::STDLINK_GUID);
        $data .= pack("H*", "02000000");
        $data .= pack("V", $options);

        return $data;
    }

    protected function getTextMarkData($textMark)
    {
        $data = pack("V", floor(strlen($textMark) / 2));
        $data .= $textMark;

        return $data;
    }
}
