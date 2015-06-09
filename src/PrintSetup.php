<?php

namespace Xls;

/**
 * Class PrintSetup
 * Contains print-related functionality for each Worksheet
 *
 * @package Xls
 */
class PrintSetup
{
    const ORIENTATION_PORTRAIT = 1;
    const ORIENTATION_LANDSCAPE = 0;

    const PAPER_CUSTOM = 0;
    const PAPER_US_LETTER = 1;
    const PAPER_A3 = 8;
    const PAPER_A4 = 9;
    const PAPER_A5 = 11;

    /**
     * The paper size
     *
     * @var integer
     */
    protected $paperSize = self::PAPER_CUSTOM;

    /**
     * Bit specifying paper orientation (for printing). 0 => landscape, 1 => portrait
     *
     * @var integer
     */
    protected $orientation = self::ORIENTATION_PORTRAIT;

    /**
     * @var Range
     */
    protected $printRepeat;

    /**
     * @var null|Range
     */
    protected $printArea = null;

    /**
     * @var float
     */
    protected $printScale = 100;

    /**
     * Whether to fit to page when printing or not.
     *
     * @var bool
     */
    protected $fitPage = false;

    /**
     * Number of pages to fit wide
     *
     * @var integer
     */
    protected $fitWidth = 0;

    /**
     * Number of pages to fit high
     *
     * @var integer
     */
    protected $fitHeight = 0;

    /**
     * The page header caption
     *
     * @var string
     */
    protected $header = '';

    /**
     * The page footer caption
     *
     * @var string
     */
    protected $footer = '';

    /**
     * The horizontal centering value for the page
     *
     * @var bool
     */
    protected $hcenter = false;

    /**
     * The vertical centering value for the page
     *
     * @var bool
     */
    protected $vcenter = false;

    /**
     * @var Margin
     */
    protected $margin;

    protected $printRowColHeaders = false;
    protected $hbreaks = array();
    protected $vbreaks = array();
    protected $printGridLines = true;

    /**
     *
     */
    public function __construct()
    {
        $this->margin = new Margin(0.75, 0.75, 1.00, 1.00);
        $this->margin->setHead(0.5)->setFoot(0.5);

        $this->printRepeat = new Range(null, null);
    }

    /**
     * @return bool
     */
    public function isPrintAreaSet()
    {
        return is_object($this->printArea);
    }

    /**
     * @return null|Range
     */
    public function getPrintArea()
    {
        return $this->printArea;
    }

    /**
     * Set the area of each worksheet that will be printed.
     *
     * @param integer $firstRow First row of the area to print
     * @param integer $firstCol First column of the area to print
     * @param integer $lastRow  Last row of the area to print
     * @param integer $lastCol  Last column of the area to print
     *
     * @return PrintSetup
     */
    public function setPrintArea($firstRow, $firstCol, $lastRow, $lastCol)
    {
        $this->printArea = new Range($firstRow, $firstCol, $lastRow, $lastCol);

        return $this;
    }

    /**
     * Set the rows to repeat at the top of each printed page.
     *
     * @param integer $firstRow First row to repeat
     * @param integer $lastRow  Last row to repeat. Optional.
     *
     * @return PrintSetup
     */
    public function printRepeatRows($firstRow, $lastRow = null)
    {
        if (!isset($lastRow)) {
            $lastRow = $firstRow;
        }

        $this->printRepeat
            ->setRowFrom($firstRow)
            ->setRowTo($lastRow)
        ;

        if (is_null($this->printRepeat->getColFrom())) {
            $this->printRepeat
                ->setColFrom(0)
                ->setColTo(Biff8::MAX_COL_IDX)
            ;
        }

        return $this;
    }

    /**
     * Set the columns to repeat at the left hand side of each printed page.
     *
     * @param integer $firstCol First column to repeat
     * @param integer $lastCol  Last column to repeat. Optional.
     *
     * @return PrintSetup
     */
    public function printRepeatColumns($firstCol, $lastCol = null)
    {
        if (!isset($lastCol)) {
            $lastCol = $firstCol;
        }

        $this->printRepeat
            ->setColFrom($firstCol)
            ->setColTo($lastCol)
        ;

        if (is_null($this->printRepeat->getRowFrom())) {
            $this->printRepeat
                ->setRowFrom(0)
                ->setRowTo(Biff8::MAX_ROW_IDX)
            ;
        }

        return $this;
    }

    /**
     * @return Range
     */
    public function getPrintRepeat()
    {
        return $this->printRepeat;
    }

    /**
     * Set the scale factor for the printed page.
     * It turns off the "fit to page" option
     *
     * @param integer $scale The optional scale factor. Defaults to 100
     *
     * @throws \Exception
     * @return PrintSetup
     */
    public function setPrintScale($scale = 100)
    {
        // Confine the scale to Excel's range
        if ($scale < 10 || $scale > 400) {
            throw new \Exception("Print scale $scale outside range: 10 <= scale <= 400");
        }

        // Turn off "fit to page" option
        $this->fitPage = false;

        $this->printScale = floor($scale);

        return $this;
    }

    /**
     * @return float
     */
    public function getPrintScale()
    {
        return $this->printScale;
    }

    /**
     * Set the paper type
     *
     * @param integer $size The type of paper size to use
     *
     * @return PrintSetup
     */
    public function setPaper($size = self::PAPER_CUSTOM)
    {
        $this->paperSize = $size;

        return $this;
    }

    /**
     * Set the page orientation as portrait.
     *
     * @return PrintSetup
     */
    public function setPortrait()
    {
        $this->orientation = self::ORIENTATION_PORTRAIT;

        return $this;
    }

    /**
     * Set the page orientation as landscape.
     *
     * @return PrintSetup
     */
    public function setLandscape()
    {
        $this->orientation = self::ORIENTATION_LANDSCAPE;

        return $this;
    }

    /**
     * @return int
     */
    public function getOrientation()
    {
        return $this->orientation;
    }

    /**
     * @return int
     */
    public function getPaperSize()
    {
        return $this->paperSize;
    }

    /**
     * @return int
     */
    public function getFitWidth()
    {
        return $this->fitWidth;
    }

    /**
     * @return int
     */
    public function getFitHeight()
    {
        return $this->fitHeight;
    }

    /**
     * @return boolean
     */
    public function isFitPage()
    {
        return $this->fitPage;
    }

    /**
     * Set the vertical and horizontal number of pages that will define the maximum area printed.
     * It doesn't seem to work with OpenOffice.
     *
     * @param  integer $width  Maximun width of printed area in pages
     * @param  integer $height Maximun heigth of printed area in pages
     *
     * @return PrintSetup
     */
    public function fitToPages($width, $height)
    {
        $this->fitPage = true;
        $this->fitWidth = $width;
        $this->fitHeight = $height;

        return $this;
    }

    /**
     * Set the page header caption and optional margin.
     *
     * @param string $string The header text
     * @param float  $margin optional head margin in inches.
     *
     * @return PrintSetup
     */
    public function setHeader($string, $margin = 0.50)
    {
        $this->header = $this->truncateStringIfNeeded($string);
        $this->margin->setHead($margin);

        return $this;
    }

    /**
     * Set the page footer caption and optional margin.
     *
     * @param string $string The footer text
     * @param float  $margin optional foot margin in inches.
     *
     * @return PrintSetup
     */
    public function setFooter($string, $margin = 0.50)
    {
        $this->footer = $this->truncateStringIfNeeded($string);
        $this->margin->setFoot($margin);

        return $this;
    }

    /**
     * @param $string
     *
     * @return string
     */
    protected function truncateStringIfNeeded($string)
    {
        if (StringUtils::countCharacters($string) > Biff8::MAX_STR_LENGTH) {
            $string = StringUtils::substr($string, 0, Biff8::MAX_STR_LENGTH);
        }

        return $string;
    }

    /**
     * Center the page horinzontally.
     *
     * @param bool $enable the optional value for centering. Defaults to 1 (center).
     *
     * @return PrintSetup
     */
    public function centerHorizontally($enable = true)
    {
        $this->hcenter = $enable;

        return $this;
    }

    /**
     * Center the page vertically.
     *
     * @param bool $enable the optional value for centering. Defaults to 1 (center).
     *
     * @return PrintSetup
     */
    public function centerVertically($enable = true)
    {
        $this->vcenter = $enable;

        return $this;
    }

    /**
     * Set the option to print the row and column headers on the printed page.
     *
     * @param bool $print Whether to print the headers or not. Defaults to 1 (print).
     *
     * @return PrintSetup
     */
    public function printRowColHeaders($print = true)
    {
        $this->printRowColHeaders = $print;

        return $this;
    }

    /**
     * Store the horizontal page breaks on a worksheet (for printing).
     * The breaks represent the row after which the break is inserted.
     *
     * @param array $breaks Array containing the horizontal page breaks
     *
     * @return PrintSetup
     */
    public function setHPagebreaks($breaks)
    {
        foreach ($breaks as $break) {
            array_push($this->hbreaks, $break);
        }

        return $this;
    }

    /**
     * Store the vertical page breaks on a worksheet (for printing).
     * The breaks represent the column after which the break is inserted.
     *
     * @param array $breaks Array containing the vertical page breaks
     *
     * @return PrintSetup
     */
    public function setVPagebreaks($breaks)
    {
        foreach ($breaks as $break) {
            array_push($this->vbreaks, $break);
        }

        return $this;
    }

    /**
     * Set the option to hide gridlines on the printed page.
     *
     * @param bool $enable
     *
     * @return PrintSetup
     */
    public function printGridlines($enable = true)
    {
        $this->printGridLines = $enable;

        return $this;
    }

    /**
     * @return bool
     */
    public function shouldPrintRowColHeaders()
    {
        return (bool)$this->printRowColHeaders;
    }

    /**
     * @return array
     */
    public function getHbreaks()
    {
        return $this->hbreaks;
    }

    /**
     * @return array
     */
    public function getVbreaks()
    {
        return $this->vbreaks;
    }

    /**
     * @return string
     */
    public function getHeader()
    {
        return $this->header;
    }

    /**
     * @return string
     */
    public function getFooter()
    {
        return $this->footer;
    }

    /**
     * @return bool
     */
    public function isHcenteringOn()
    {
        return (bool)$this->hcenter;
    }

    /**
     * @return bool
     */
    public function isVcenteringOn()
    {
        return (bool)$this->vcenter;
    }

    /**
     * @return bool
     */
    public function shouldPrintGridLines()
    {
        return (bool)$this->printGridLines;
    }

    /**
     * @return Margin
     */
    public function getMargin()
    {
        return $this->margin;
    }
}
