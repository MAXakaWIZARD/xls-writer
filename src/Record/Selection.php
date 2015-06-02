<?php

namespace Xls\Record;

class Selection extends AbstractRecord
{
    const NAME = 'SELECTION';
    const ID = 0x001D;

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

        return $this->getFullRecord($data);
    }
}
