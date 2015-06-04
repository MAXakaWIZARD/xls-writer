<?php

namespace Xls;

class Range
{
    protected $colFrom;
    protected $colTo;
    protected $rowFrom;
    protected $rowTo;

    /**
     * @param int  $rowFrom
     * @param int  $colFrom
     * @param int $rowTo
     * @param int $colTo
     * @param bool $normalize
     */
    public function __construct(
        $rowFrom = 0,
        $colFrom = 0,
        $rowTo = null,
        $colTo = null,
        $normalize = true
    ) {
        if (!isset($rowTo)) {
            $rowTo = $rowFrom; // Last row in reference
        }

        if (!isset($colTo)) {
            $colTo = $colFrom; // Last col in reference
        }

        $this->colFrom = $colFrom;
        $this->colTo = $colTo;
        $this->rowFrom = $rowFrom;
        $this->rowTo = $rowTo;

        if ($normalize) {
            $this->normalize();
        }
    }

    /**
     * Swap last row/col for first row/col as necessary
     */
    protected function normalize()
    {
        if ($this->rowFrom > $this->rowTo) {
            $tmp = $this->rowFrom;
            $this->rowTo = $this->rowFrom;
            $this->rowFrom = $tmp;
        }

        if ($this->colFrom > $this->colTo) {
            $tmp = $this->colFrom;
            $this->colTo = $this->colFrom;
            $this->colFrom = $tmp;
        }
    }

    /**
     * @return integer
     */
    public function getColFrom()
    {
        return $this->colFrom;
    }

    /**
     * @return integer
     */
    public function getColTo()
    {
        return $this->colTo;
    }

    /**
     * @return integer
     */
    public function getRowFrom()
    {
        return $this->rowFrom;
    }

    /**
     * @return integer
     */
    public function getRowTo()
    {
        return $this->rowTo;
    }

    /**
     * Include specified cell
     * @param Cell $cell
     */
    public function expand(Cell $cell)
    {
        if ($cell->getRow() < $this->rowFrom) {
            $this->rowFrom = $cell->getRow();
        }

        if ($cell->getRow() > $this->rowTo) {
            $this->rowTo = $cell->getRow();
        }

        if ($cell->getCol() < $this->colFrom) {
            $this->colFrom = $cell->getCol();
        }

        if ($cell->getCol() > $this->colTo) {
            $this->colTo = $cell->getCol();
        }
    }

    /**
     * @return Cell
     */
    public function startCell()
    {
        return new Cell($this->rowFrom, $this->colFrom);
    }
}
