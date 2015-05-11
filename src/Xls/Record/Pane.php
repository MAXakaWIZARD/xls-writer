<?php
namespace Xls\Record;

class Pane extends AbstractRecord
{
    const NAME = 'PANE';
    const ID = 0x0041;
    const LENGTH = 0x0A;

    /**
     * @param $x
     * @param $y
     * @param $rowTop
     * @param $colLeft
     * @param $activePane
     *
     * @return string
     */
    public function getData($x, $y, $rowTop, $colLeft, $activePane)
    {
        $data = pack("vvvvv", $x, $y, $rowTop, $colLeft, $activePane);

        return $this->getHeader() . $data;
    }
}
