<?php
namespace Xls\Record;

class Hyperlink extends AbstractRecord
{
    const NAME = 'HYPERLINK';
    const ID = 0x01B8;

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
        $options = pack("V", 0x03);

        // Convert URL to a null terminated wchar string
        $url = join("\0", preg_split("''", $url, -1, PREG_SPLIT_NO_EMPTY));
        $url = $url . "\0\0\0";

        // Pack the length of the URL
        $urlLen = pack("V", strlen($url));

        // Calculate the data length
        $length = 0x34 + strlen($url);

        // Pack the header data
        $data = pack("vvvv", $row1, $row2, $col1, $col2);

        // Pack the undocumented parts of the hyperlink stream
        $data .= pack("H*", "D0C9EA79F9BACE118C8200AA004BA90B02000000");
        $data .= $options;
        $data .= pack("H*", "E0C9EA79F9BACE118C8200AA004BA90B");
        $data .= $urlLen . $url;


        return $this->getHeader($length) . $data;
    }
}
