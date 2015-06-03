<?php

namespace Xls\Record;

use Xls\Range;

class Selection extends AbstractRecord
{
    const NAME = 'SELECTION';
    const ID = 0x001D;

    /**
     * Generate the SELECTION record
     *
     * @param Range $selection
     * @param integer $activePane pane position
     * @return string
     */
    public function getData($selection, $activePane)
    {
        $rwAct = $selection->getRowFrom(); // Active row
        $colAct = $selection->getColFrom(); // Active column
        $irefAct = 0; // Active cell ref
        $cref = 1; // Number of refs

        $data = pack(
            "CvvvvvvCC",
            $activePane,
            $rwAct,
            $colAct,
            $irefAct,
            $cref,
            $selection->getRowFrom(),
            $selection->getRowTo(),
            $selection->getColFrom(),
            $selection->getColTo()
        );

        return $this->getFullRecord($data);
    }
}
