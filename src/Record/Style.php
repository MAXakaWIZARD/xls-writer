<?php

namespace Xls\Record;

class Style extends AbstractRecord
{
    const NAME = 'STYLE';
    const ID = 0x0293;

    /**
     * Generate the TYLE records
     * @return string
     */
    public function getData()
    {
        $ixfe = 0x8000; // Index to style XF
        $builtIn = 0x00; // Built-in style
        $iLevel = 0xff; // Outline style level

        $data = pack("vCC", $ixfe, $builtIn, $iLevel);

        return $this->getFullRecord($data);
    }
}
