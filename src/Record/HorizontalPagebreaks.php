<?php
namespace Xls\Record;

class HorizontalPagebreaks extends AbstractRecord
{
    const NAME = 'HORIZONTALPAGEBREAKS';
    const ID = 0x001B;
    const COUNT_LIMIT = 0;

    /**
     * @param array $breaks
     *
     * @return string
     */
    public function getData($breaks)
    {
        if (static::COUNT_LIMIT > 0) {
            $breaks = array_slice($breaks, 0, static::COUNT_LIMIT);
        }

        sort($breaks, SORT_NUMERIC);
        if ($breaks[0] == 0) {
            // don't use first break if it's 0
            array_shift($breaks);
        }

        $cbrk = count($breaks);
        $data = pack("v", $cbrk);

        // Append each page break
        foreach ($breaks as $break) {
            $data .= pack("vvv", $break, 0x00, 0xff);
        }

        return $this->getFullRecord($data);
    }
}
