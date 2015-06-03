<?php
namespace Xls\Record;

use Xls\StringUtils;
use Xls\Range;

class HyperlinkExternal extends Hyperlink
{
    const MONIKER_GUID = "0303000000000000C000000000000046";

    /**
     * @param Range $range
     * @param $url
     *
     * @return string
     */
    public function getData(Range $range, $url)
    {
        $cellRef = null;
        if (preg_match("/\#/", $url)) {
            $parts = explode('#', $url);
            $url = $parts[0];
            $cellRef = $parts[1];
        }

        // Calculate the up-level dir count e.g.. (..\..\..\ == 3)
        $upCount = preg_match_all("/\.\.\\\/", $url, $useless);

        // Store the short dos dir name (null terminated)
        $urlWithoutDots = preg_replace("/\.\.\\\/", '', $url) . "\0";

        // Unknown 24 bytes
        $unknown = pack("H*", 'FFFFADDE' . str_repeat('00', 20));

        $streamLen = pack("V", 0);

        $options = $this->getOptions($url);
        $data = $this->getCommonData($range, $options);
        $data .= pack("H*", static::MONIKER_GUID) .
            pack("v", $upCount) .
            pack("V", strlen($urlWithoutDots)) .
            $urlWithoutDots .
            $unknown .
            $streamLen;

        if ($cellRef) {
            $cellRef = StringUtils::toNullTerminatedWchar($cellRef);
            $data .= $this->getTextMarkData($cellRef);
        }

        return $this->getFullRecord($data);
    }

    protected function getOptions($url)
    {
        // Determine if the link is relative or absolute:
        //   relative if link contains no dir separator, "somefile.xls"
        //   relative if link starts with up-dir, "..\..\somefile.xls"
        //   otherwise, absolute

        $absolute = 1; // Bit mask
        if (!preg_match("/\\\/", $url)
            || preg_match("/^\.\.\\\/", $url)
        ) {
            $absolute = 0;
        }

        $options = 0x00;
        $options |= 1 << 0; //File link or URL
        $options |= $absolute << 1; //File link or URL

        if (preg_match("/\#/", $url)) {
            $options |= 1 << 3; //Has text mark
        }

        return $options;
    }
}
