<?php

namespace Xls\Record;

class Selection extends AbstractRecord
{
    const NAME = 'SELECTION';
    const ID = 0x001D;
    const LENGTH = 0x0F;

    /**
     * Generate the SELECTION record
     *
     * @param array $selection array containing ($rwFirst,$colFirst,$rwLast,$colLast)
     * @param integer $activePane pane position
     * @return string
     */
    public function getData($selection, $activePane)
    {
        list($rwFirst, $colFirst, $rwLast, $colLast) = $selection;

        $rwAct = $rwFirst; // Active row
        $colAct = $colFirst; // Active column
        $irefAct = 0; // Active cell ref
        $cref = 1; // Number of refs

        if (!isset($rwLast)) {
            $rwLast = $rwFirst; // Last row in reference
        }
        if (!isset($colLast)) {
            $colLast = $colFirst; // Last col in reference
        }

        // Swap last row/col for first row/col as necessary
        if ($rwFirst > $rwLast) {
            list($rwFirst, $rwLast) = array($rwLast, $rwFirst);
        }

        if ($colFirst > $colLast) {
            list($colFirst, $colLast) = array($colLast, $colFirst);
        }

        $data = pack(
            "CvvvvvvCC",
            $activePane,
            $rwAct,
            $colAct,
            $irefAct,
            $cref,
            $rwFirst,
            $rwLast,
            $colFirst,
            $colLast
        );

        return $this->getHeader() . $data;
    }
}
