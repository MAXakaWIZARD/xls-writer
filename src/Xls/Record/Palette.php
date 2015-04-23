<?php

namespace Xls\Record;

class Palette extends AbstractRecord
{
    const NAME = 'PALETTE';
    const ID = 0x0092;
    const LENGTH = 0x02;

    /**
     * Write the PALETTE biff record
     *
     * @param array $palette array with colors
     * @return string
     */
    public function getData($palette)
    {
        $ccv = count($palette); // Number of RGB values to follow
        $extraLength = 4 * $ccv; // Number of bytes to follow

        $data = ''; // The RGB data
        // Pack the RGB data
        foreach ($palette as $color) {
            foreach ($color as $byte) {
                $data .= pack("C", $byte);
            }
        }

        return $this->getHeader($extraLength, $ccv) . $data;
    }
}