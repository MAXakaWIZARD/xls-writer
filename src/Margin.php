<?php

namespace Xls;

class Margin
{
    protected $top;
    protected $bottom;
    protected $left;
    protected $right;
    protected $head;
    protected $foot;

    /**
     * @param float $left
     * @param float $right
     * @param float $top
     * @param float $bottom
     */
    public function __construct($left, $right, $top, $bottom)
    {
        $this->top = $top;
        $this->bottom = $bottom;
        $this->left = $left;
        $this->right = $right;
    }

    /**
     * @return float
     */
    public function getTop()
    {
        return $this->top;
    }

    /**
     * @param float $top
     * @return Margin
     */
    public function setTop($top)
    {
        $this->top = $top;

        return $this;
    }

    /**
     * @return float
     */
    public function getBottom()
    {
        return $this->bottom;
    }

    /**
     * @param float $bottom
     * @return Margin
     */
    public function setBottom($bottom)
    {
        $this->bottom = $bottom;

        return $this;
    }

    /**
     * @return float
     */
    public function getLeft()
    {
        return $this->left;
    }

    /**
     * @param float $left
     * @return Margin
     */
    public function setLeft($left)
    {
        $this->left = $left;

        return $this;
    }

    /**
     * @return float
     */
    public function getRight()
    {
        return $this->right;
    }

    /**
     * @param float $right
     * @return Margin
     */
    public function setRight($right)
    {
        $this->right = $right;

        return $this;
    }

    /**
     * @return float
     */
    public function getHead()
    {
        return $this->head;
    }

    /**
     * @param float $head
     * @return Margin
     */
    public function setHead($head)
    {
        $this->head = $head;

        return $this;
    }

    /**
     * @return float
     */
    public function getFoot()
    {
        return $this->foot;
    }

    /**
     * @param float $foot
     * @return Margin
     */
    public function setFoot($foot)
    {
        $this->foot = $foot;

        return $this;
    }

    /**
     * @param float $margin
     * @return Margin
     */
    public function setAll($margin)
    {
        $this->setTop($margin);
        $this->setBottom($margin);
        $this->setLeft($margin);
        $this->setRight($margin);

        return $this;
    }
}
