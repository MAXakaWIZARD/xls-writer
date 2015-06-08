<?php

namespace Xls;

class PageSetup
{
    const ORIENTATION_PORTRAIT = 1;
    const ORIENTATION_LANDSCAPE = 0;

    const PAPER_CUSTOM = 0;
    const PAPER_US_LETTER = 1;
    const PAPER_A3 = 8;
    const PAPER_A4 = 9;
    const PAPER_A5 = 11;

    /**
     * The paper size (for printing) (DOCUMENT!!!)
     * @var integer
     */
    protected $paperSize = self::PAPER_CUSTOM;

    /**
     * Bit specifying paper orientation (for printing). 0 => landscape, 1 => portrait
     * @var integer
     */
    protected $orientation = self::ORIENTATION_PORTRAIT;

    /**
     * First row to reapeat on each printed page
     * @var integer
     */
    protected $titleRowMin = null;

    /**
     * Last row to reapeat on each printed page
     * @var integer
     */
    protected $titleRowMax = null;

    /**
     * First column to reapeat on each printed page
     * @var integer
     */
    protected $titleColMin = null;

    /**
     * Last column to reapeat on each printed page
     * @var integer
     */
    protected $titleColMax = null;

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
     * @var bool
     */
    protected $fitPage = false;

    /**
     * Number of pages to fit wide
     * @var integer
     */
    protected $fitWidth = 0;

    /**
     * Number of pages to fit high
     * @var integer
     */
    protected $fitHeight = 0;

    /**
     * The page header caption
     * @var string
     */
    protected $header = '';

    /**
     * The page footer caption
     * @var string
     */
    protected $footer = '';

    /**
     * The horizontal centering value for the page
     * @var bool
     */
    protected $hcenter = false;

    /**
     * The vertical centering value for the page
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
    protected $screenGridLines = true;

    /**
     * @var float
     */
    protected $zoom = 100;

    /**
     *
     */
    public function __construct()
    {
        $this->margin = new Margin(0.75, 0.75, 1.00, 1.00);
        $this->margin->setHead(0.5)->setFoot(0.5);
    }

    /**
     * @return bool
     */
    public function isPrintAreaSet()
    {
        return !is_null($this->printArea);
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
     * @param integer $firstRow First row of the area to print
     * @param integer $firstCol First column of the area to print
     * @param integer $lastRow  Last row of the area to print
     * @param integer $lastCol  Last column of the area to print
     * @return PageSetup
     */
    public function setPrintArea($firstRow, $firstCol, $lastRow, $lastCol)
    {
        $this->printArea = new Range($firstRow, $firstCol, $lastRow, $lastCol);

        return $this;
    }

    /**
     * Set the rows to repeat at the top of each printed page.
     * @param integer $firstRow First row to repeat
     * @param integer $lastRow  Last row to repeat. Optional.
     * @return PageSetup
     */
    public function printRepeatRows($firstRow, $lastRow = null)
    {
        if (!isset($lastRow)) {
            $lastRow = $firstRow;
        }

        $this->titleRowMin = $firstRow;
        $this->titleRowMax = $lastRow;

        return $this;
    }

    /**
     * Set the columns to repeat at the left hand side of each printed page.
     * @param integer $firstCol First column to repeat
     * @param integer $lastCol  Last column to repeat. Optional.
     * @return PageSetup
     */
    public function printRepeatColumns($firstCol, $lastCol = null)
    {
        if (!isset($lastCol)) {
            $lastCol = $firstCol;
        }

        $this->titleColMin = $firstCol;
        $this->titleColMax = $lastCol;

        return $this;
    }

    /**
     * @return int
     */
    public function getTitleRowMin()
    {
        return $this->titleRowMin;
    }

    /**
     * @return int
     */
    public function getTitleRowMax()
    {
        return $this->titleRowMax;
    }

    /**
     * @return int
     */
    public function getTitleColMin()
    {
        return $this->titleColMin;
    }

    /**
     * @return int
     */
    public function getTitleColMax()
    {
        return $this->titleColMax;
    }

    /**
     * Set the scale factor for the printed page.
     * It turns off the "fit to page" option
     * @param integer $scale The optional scale factor. Defaults to 100
     * @throws \Exception
     * @return PageSetup
     */
    public function setPrintScale($scale = 100)
    {
        // Confine the scale to Excel's range
        if ($scale < 10 || $scale > 400) {
            throw new \Exception("Print scale $scale outside range: 10 <= zoom <= 400");
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
     * @param integer $size The type of paper size to use
     * @return PageSetup
     */
    public function setPaper($size = self::PAPER_CUSTOM)
    {
        $this->paperSize = $size;

        return $this;
    }

    /**
     * Set the page orientation as portrait.
     * @return PageSetup
     */
    public function setPortrait()
    {
        $this->orientation = self::ORIENTATION_PORTRAIT;

        return $this;
    }

    /**
     * Set the page orientation as landscape.
     * @return PageSetup
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
     * @param  integer $width  Maximun width of printed area in pages
     * @param  integer $height Maximun heigth of printed area in pages
     * @return PageSetup
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
     * @param string $string The header text
     * @param float $margin optional head margin in inches.
     * @return PageSetup
     */
    public function setHeader($string, $margin = 0.50)
    {
        $this->header = $this->truncateStringIfNeeded($string);
        $this->margin->setHead($margin);

        return $this;
    }

    /**
     * Set the page footer caption and optional margin.
     * @param string $string The footer text
     * @param float $margin optional foot margin in inches.
     * @return PageSetup
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
     * @return PageSetup
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
     * @return PageSetup
     */
    public function centerVertically($enable = true)
    {
        $this->vcenter = $enable;

        return $this;
    }

    /**
     * Set the option to print the row and column headers on the printed page.
     * @param bool $print Whether to print the headers or not. Defaults to 1 (print).
     * @return PageSetup
     */
    public function printRowColHeaders($print = true)
    {
        $this->printRowColHeaders = $print;

        return $this;
    }

    /**
     * Store the horizontal page breaks on a worksheet (for printing).
     * The breaks represent the row after which the break is inserted.
     * @param array $breaks Array containing the horizontal page breaks
     * @return PageSetup
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
     * @param array $breaks Array containing the vertical page breaks
     * @return PageSetup
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
     * @return PageSetup
     */
    public function printGridlines($enable = true)
    {
        $this->printGridLines = $enable;

        return $this;
    }

    /**
     * Set the option to hide gridlines on the worksheet (as seen on the screen).
     * @param bool $visible
     * @return PageSetup
     */
    public function showGridlines($visible = true)
    {
        $this->screenGridLines = $visible;

        return $this;
    }

    /**
     * Set the worksheet zoom factor.
     *
     * @param integer $percents The zoom factor
     *
     * @throws \Exception
     * @return PageSetup
     */
    public function setZoom($percents = 100)
    {
        // Confine the scale to Excel's range
        if ($percents < 10 || $percents > 400) {
            throw new \Exception("Zoom factor $percents outside range: 10 <= zoom <= 400");
        }

        $this->zoom = floor($percents);

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
    public function areGridLinesVisible()
    {
        return (bool)$this->screenGridLines;
    }

    /**
     * @return bool
     */
    public function shouldPrintGridLines()
    {
        return (bool)$this->printGridLines;
    }

    /**
     * @return float
     */
    public function getZoom()
    {
        return $this->zoom;
    }

    /**
     * @return Margin
     */
    public function getMargin()
    {
        return $this->margin;
    }
}
