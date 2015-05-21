<?php
namespace Xls\Record;

use Xls\StringUtils;

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

        if (StringUtils::isIconvEnabled() && StringUtils::isMbstringEnabled()) {
            $format = StringUtils::toUtf16Le($format);
            $encoding = 1;
        } else {
            $encoding = 0;
        }

        $cch = StringUtils::countCharacters($format);

        $data = pack("vvC", $formatIndex, $cch, $encoding);

        return $this->getHeader($length) . $data . $format;
    }
}
