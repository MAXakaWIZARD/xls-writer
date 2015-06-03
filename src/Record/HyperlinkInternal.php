<?php
namespace Xls\Record;

use Xls\StringUtils;
use Xls\Range;

class HyperlinkInternal extends Hyperlink
{
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
        $data .= $this->getTextMarkData($url);

        return $this->getFullRecord($data);
    }

    protected function getOptions($url)
    {
        $options = 0x00;
        $options |= 1 << 3; //Has Text mark

        return $options;
    }
}
