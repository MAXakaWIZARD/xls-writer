<?php

namespace Xls;

class Cell
{
    /**
     * Utility function for writing formulas
     * Converts a cell's coordinates to the A1 format.
     *
     * @param integer $row Row for the cell to convert (0-indexed).
     * @param integer $col Column for the cell to convert (0-indexed).
     *
     * @throws \Exception
     * @return string The cell identifier in A1 format
     */
    public static function getAddress($row, $col)
    {
        if ($col >= Biff5::MAX_COLS) {
            throw new \Exception("Maximum column value exceeded: $col");
        }

        $int = (int)($col / 26);
        $frac = $col % 26;
        $chr1 = '';

        if ($int > 0) {
            $chr1 = chr(ord('A') + $int - 1);
        }

        $chr2 = chr(ord('A') + $frac);
        $row++;

        return $chr1 . $chr2 . $row;
    }
}
