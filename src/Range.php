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
     * @return Range
     */
    public function expand(Cell $cell)
    {
        if ($cell->getRow() > $this->rowTo) {
            $this->setRowTo($cell->getRow());
        }

        if ($cell->getCol() > $this->colTo) {
            $this->setColTo($cell->getCol());
        }

        return $this;
    }

    /**
     * @return Cell
     */
    public function getStartCell()
    {
        return new Cell($this->rowFrom, $this->colFrom);
    }

    /**
     * @param int $colFrom
     * @return Range
     */
    public function setColFrom($colFrom)
    {
        $this->colFrom = $colFrom;

        return $this;
    }

    /**
     * @param int|null $colTo
     * @return Range
     */
    public function setColTo($colTo)
    {
        $this->colTo = $colTo;

        return $this;
    }

    /**
     * @param int $rowFrom
     * @return Range
     */
    public function setRowFrom($rowFrom)
    {
        $this->rowFrom = $rowFrom;

        return $this;
    }

    /**
     * @param int|null $rowTo
     * @return Range
     */
    public function setRowTo($rowTo)
    {
        $this->rowTo = $rowTo;

        return $this;
    }

    /**
     * @return bool
     */
    public function isEmpty()
    {
        return is_null($this->getRowFrom()) && is_null($this->getColFrom());
    }
}
