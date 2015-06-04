<?php

namespace Xls;

class Cell
{
    protected $row;
    protected $col;

    public function __construct($row, $col)
    {
        $this->validateRowIndex($row);
        $this->validateColIndex($col);

        $this->row = $row;
        $this->col = $col;
    }

    /**
     * @return mixed
     */
    public function getRow()
    {
        return $this->row;
    }

    /**
     * @return mixed
     */
    public function getCol()
    {
        return $this->col;
    }

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
        if ($col >= Biff8::MAX_COLS) {
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

    /**
     * Convert an Excel cell reference such as A1 or $B2 or C$3 or $D$4 to a zero
     * indexed row and column number. Also returns two (0,1) values to indicate
     * whether the row or column are relative references.
     *
     * @param string $address The Excel cell reference in A1 format.
     * @return array
     */
    public static function addressToRowCol($address)
    {
        preg_match('/(\$)?([A-Z]+)(\$)?(\d+)/', $address, $match);
        // return absolute column if there is a $ in the ref
        $colRel = empty($match[1]) ? 1 : 0;
        $colRef = $match[2];
        $rowRel = empty($match[3]) ? 1 : 0;
        $row = $match[4];

        // Convert base26 column string to a number.
        $expn = strlen($colRef) - 1;
        $col = 0;
        $colRefLength = strlen($colRef);
        for ($i = 0; $i < $colRefLength; $i++) {
            $col += (ord($colRef{$i}) - ord('A') + 1) * pow(26, $expn);
            $expn--;
        }

        // Convert 1-index to zero-index
        $row--;
        $col--;

        return array($row, $col, $rowRel, $colRel);
    }

    /**
     * @param $row
     *
     * @throws \Exception
     */
    protected function validateRowIndex($row)
    {
        if ($row >= Biff8::MAX_ROWS) {
            throw new \Exception('Row index is beyond max row number');
        }
    }

    /**
     * @param $col
     *
     * @throws \Exception
     */
    protected function validateColIndex($col)
    {
        if ($col >= Biff8::MAX_COLS) {
            throw new \Exception('Col index is beyond max col number');
        }
    }
}
