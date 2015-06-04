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
    protected $fitPage = 0;

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
     * @var integer
     */
    protected $hcenter = 0;

    /**
     * The vertical centering value for the page
     * @var integer
     */
    protected $vcenter = 0;

    /**
     * The margin for the header
     * @var float
     */
    protected $marginHead = 0.50;

    /**
     * The margin for the footer
     * @var float
     */
    protected $marginFoot = 0.50;

    /**
     * The left margin for the worksheet in inches
     * @var float
     */
    protected $marginLeft = 0.75;

    /**
     * The right margin for the worksheet in inches
     * @var float
     */
    protected $marginRight = 0.75;

    /**
     * The top margin for the worksheet in inches
     * @var float
     */
    protected $marginTop = 1.00;

    /**
     * The bottom margin for the worksheet in inches
     * @var float
     */
    protected $marginBottom = 1.00;

    protected $printRowColHeaders = 0;
    protected $hbreaks = array();
    protected $vbreaks = array();
    protected $printGridLines = 1;
    protected $screenGridLines = 1;

    /**
     * @var float
     */
    protected $zoom = 100;

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
     */
    public function setPrintArea($firstRow, $firstCol, $lastRow, $lastCol)
    {
        $this->printArea = new Range($firstRow, $firstCol, $lastRow, $lastCol);
    }

    /**
     * Set the rows to repeat at the top of each printed page.
     * @param integer $firstRow First row to repeat
     * @param integer $lastRow  Last row to repeat. Optional.
     */
    public function printRepeatRows($firstRow, $lastRow = null)
    {
        if (!isset($lastRow)) {
            $lastRow = $firstRow;
        }

        $this->titleRowMin = $firstRow;
        $this->titleRowMax = $lastRow;
    }

    /**
     * Set the columns to repeat at the left hand side of each printed page.
     * @param integer $firstCol First column to repeat
     * @param integer $lastCol  Last column to repeat. Optional.
     */
    public function printRepeatColumns($firstCol, $lastCol = null)
    {
        if (!isset($lastCol)) {
            $lastCol = $firstCol;
        }

        $this->titleColMin = $firstCol;
        $this->titleColMax = $lastCol;
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
     */
    public function setPrintScale($scale = 100)
    {
        // Confine the scale to Excel's range
        if ($scale < 10 || $scale > 400) {
            throw new \Exception("Print scale $scale outside range: 10 <= zoom <= 400");
        }

        // Turn off "fit to page" option
        $this->fitPage = 0;

        $this->printScale = floor($scale);
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
     */
    public function setPaper($size = self::PAPER_CUSTOM)
    {
        $this->paperSize = $size;
    }

    /**
     * Set the page orientation as portrait.
     */
    public function setPortrait()
    {
        $this->orientation = self::ORIENTATION_PORTRAIT;
    }

    /**
     * Set the page orientation as landscape.
     */
    public function setLandscape()
    {
        $this->orientation = self::ORIENTATION_LANDSCAPE;
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
     * @see setPrintScale()
     */
    public function fitToPages($width, $height)
    {
        $this->fitPage = 1;
        $this->fitWidth = $width;
        $this->fitHeight = $height;
    }

    /**
     * Set the page header caption and optional margin.
     * @param string $string The header text
     * @param float $margin optional head margin in inches.
     */
    public function setHeader($string, $margin = 0.50)
    {
        $this->header = $this->truncateStringIfNeeded($string);
        $this->marginHead = $margin;
    }

    /**
     * Set the page footer caption and optional margin.
     * @param string $string The footer text
     * @param float $margin optional foot margin in inches.
     */
    public function setFooter($string, $margin = 0.50)
    {
        $this->footer = $this->truncateStringIfNeeded($string);
        $this->marginFoot = $margin;
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
     * @param integer $center the optional value for centering. Defaults to 1 (center).
     */
    public function centerHorizontally($center = 1)
    {
        $this->hcenter = $center;
    }

    /**
     * Center the page vertically.
     * @param integer $center the optional value for centering. Defaults to 1 (center).
     */
    public function centerVertically($center = 1)
    {
        $this->vcenter = $center;
    }

    /**
     * Set all the page margins to the same value in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMargins($margin)
    {
        $this->setMarginsLeftRight($margin);
        $this->setMarginsTopBottom($margin);
    }

    /**
     * Set the left and right margins to the same value in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginsLeftRight($margin)
    {
        $this->setMarginLeft($margin);
        $this->setMarginRight($margin);
    }

    /**
     * Set the top and bottom margins to the same value in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginsTopBottom($margin)
    {
        $this->setMarginTop($margin);
        $this->setMarginBottom($margin);
    }

    /**
     * Set the left margin in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginLeft($margin = 0.75)
    {
        $this->marginLeft = $margin;
    }

    /**
     * Set the right margin in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginRight($margin = 0.75)
    {
        $this->marginRight = $margin;
    }

    /**
     * Set the top margin in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginTop($margin = 1.00)
    {
        $this->marginTop = $margin;
    }

    /**
     * Set the bottom margin in inches.
     * @param float $margin The margin to set in inches
     */
    public function setMarginBottom($margin = 1.00)
    {
        $this->marginBottom = $margin;
    }

    /**
     * Set the option to print the row and column headers on the printed page.
     * @param integer $print Whether to print the headers or not. Defaults to 1 (print).
     */
    public function printRowColHeaders($print = 1)
    {
        $this->printRowColHeaders = $print;
    }

    /**
     * Store the horizontal page breaks on a worksheet (for printing).
     * The breaks represent the row after which the break is inserted.
     * @param array $breaks Array containing the horizontal page breaks
     */
    public function setHPagebreaks($breaks)
    {
        foreach ($breaks as $break) {
            array_push($this->hbreaks, $break);
        }
    }

    /**
     * Store the vertical page breaks on a worksheet (for printing).
     * The breaks represent the column after which the break is inserted.
     * @param array $breaks Array containing the vertical page breaks
     */
    public function setVPagebreaks($breaks)
    {
        foreach ($breaks as $break) {
            array_push($this->vbreaks, $break);
        }
    }

    /**
     * Set the option to hide gridlines on the printed page.
     */
    public function hidePrintGridlines()
    {
        $this->printGridLines = 0;
    }

    /**
     * Set the option to hide gridlines on the worksheet (as seen on the screen).
     */
    public function hideScreenGridlines()
    {
        $this->screenGridLines = 0;
    }

    /**
     * Set the worksheet zoom factor.
     * @param integer $scale The zoom factor
     * @throws \Exception
     */
    public function setZoom($scale = 100)
    {
        // Confine the scale to Excel's range
        if ($scale < 10 || $scale > 400) {
            throw new \Exception("Zoom factor $scale outside range: 10 <= zoom <= 400");
        }

        $this->zoom = floor($scale);
    }

    /**
     * @return float
     */
    public function getMarginHead()
    {
        return $this->marginHead;
    }

    /**
     * @return float
     */
    public function getMarginFoot()
    {
        return $this->marginFoot;
    }

    /**
     * @return float
     */
    public function getMarginLeft()
    {
        return $this->marginLeft;
    }

    /**
     * @return float
     */
    public function getMarginRight()
    {
        return $this->marginRight;
    }

    /**
     * @return float
     */
    public function getMarginTop()
    {
        return $this->marginTop;
    }

    /**
     * @return float
     */
    public function getMarginBottom()
    {
        return $this->marginBottom;
    }

    /**
     * @return int
     */
    public function getPrintRowColHeaders()
    {
        return $this->printRowColHeaders;
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
     * @return int
     */
    public function getPrintGridLines()
    {
        return $this->printGridLines;
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
     * @return int
     */
    public function getHcenter()
    {
        return $this->hcenter;
    }

    /**
     * @return int
     */
    public function getVcenter()
    {
        return $this->vcenter;
    }

    /**
     * @return bool
     */
    public function areScreenGridLinesVisible()
    {
        return (bool)$this->screenGridLines;
    }

    /**
     * @return float
     */
    public function getZoom()
    {
        return $this->zoom;
    }
}
