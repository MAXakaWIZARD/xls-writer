<?php
namespace Xls\Record;

class Formula extends AbstractRecord
{
    const NAME = 'FORMULA';
    const ID = 0x0006;

    /**
     * @param $row
     * @param $col
     * @param string $formula Formula in reverse polish format
     * @param $format
     *
     * @return string
     * @throws \Exception
     */
    public function getData($row, $col, $formula, $format)
    {
        // Excel normally stores the last calculated value of the formula in $num.
        // Clearly we are not in a position to calculate this a priori. Instead
        // we set $num to zero and set the option flags in $grbit to ensure
        // automatic calculation of the formula when the file is opened.
        $num = 0x00; // Current value of formula
        $grbit = 0x03; // Option flags
        $unknown = 0x0000; // Must be zero

        $formlen = strlen($formula);

        $data = pack(
            "vvvdvVv",
            $row,
            $col,
            $this->xf($format),
            $num,
            $grbit,
            $unknown,
            $formlen
        );
        $data .= $formula;

        return $this->getFullRecord($data);
    }
}
