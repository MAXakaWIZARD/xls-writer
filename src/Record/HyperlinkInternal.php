<?php
namespace Xls\Record;

use Xls\StringUtils;

class HyperlinkInternal extends Hyperlink
{
    /**
     * @param $row1
     * @param $row2
     * @param $col1
     * @param $col2
     * @param $url
     *
     * @return string
     */
    public function getData($row1, $row2, $col1, $col2, $url)
    {
        $url = StringUtils::toNullTerminatedWchar($url);

        $options = $this->getOptions($url);
        $data = $this->getCommonData($row1, $row2, $col1, $col2, $options);
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
