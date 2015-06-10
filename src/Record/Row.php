<?php
namespace Xls\Record;

class Row extends AbstractRecord
{
    const NAME = 'ROW';
    const ID = 0x0208;

    /**
     * @param array $rowInfo
     *
     * @return string
     */
    public function getData($rowInfo)
    {
        $colMic = 0x0000; // First defined column
        $colMac = 0x0000; // Last defined column
        $irwMac = 0x0000; // Used by Excel to optimise loading
        $reserved = 0x0000; // Reserved

        $height = $rowInfo['height'];
        if (!is_null($height)) {
            $height = $height * 20; // row height
        } else {
            $height = 0xff; // default row height is 256
        }

        $level = max(0, min($rowInfo['level'], 7)); // level should be between 0 and 7

        $data = pack(
            "vvvvvvvv",
            $rowInfo['row'],
            $colMic,
            $colMac,
            $height,
            $irwMac,
            $reserved,
            $this->getGrBit($rowInfo['format'], $rowInfo['hidden'], $level),
            $this->xf($rowInfo['format'])
        );

        return $this->getFullRecord($data);
    }

    /**
     * Get the options flags. fUnsynced is used to show that the font and row
     * heights are not compatible. This is usually the case for WriteExcel.
     * The collapsed flag 0x10 doesn't seem to be used to indicate that a row
     * is collapsed. Instead it is used to indicate that the previous row is
     * collapsed. The zero height flag, 0x20, is used to collapse a row.
     *
     * @param $format
     * @param $hidden
     * @param $level
     *
     * @return int
     */
    protected function getGrBit($format, $hidden, $level)
    {
        $grbit = 0x0000;
        $grbit |= $level;

        if ($hidden) {
            $grbit |= 0x0020;
        }

        $grbit |= 0x0040; // fUnsynced

        if ($format) {
            $grbit |= 0x0080;
        }

        $grbit |= 0x0100;

        return $grbit;
    }
}
