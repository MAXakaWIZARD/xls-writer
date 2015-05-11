<?php

namespace Xls;

/**
 * Class for generating Excel XF records (formats)
 *
 * @author   Xavier Noguer <xnoguer@rezebra.com>
 * @category FileFormats
 * @package  Spreadsheet_Excel_Writer
 */

class Format
{
    const BORDER_THIN = 1;
    const BORDER_THICK = 2;

    const UNDERLINE_ONCE = 1;
    const UNDERLINE_TWICE = 2;

    const SCRIPT_SUPER = 1;
    const SCRIPT_SUB = 2;

    const FONT_NORMAL = 400;
    const FONT_BOLD = 700;

    /**
     * @var integer
     */
    protected $version;

    /**
     * The byte order of this architecture. 0 => little endian, 1 => big endian
     * @var integer
     */
    protected $byteOrder;

    /**
     * The index given by the workbook when creating a new format.
     * @var integer
     */
    public $xfIndex;

    /**
     * Index to the FONT record.
     * @var integer
     */
    public $fontIndex = 0;

    /**
     * The font name (ASCII).
     * @var string
     */
    public $fontName = 'Arial';

    /**
     * Height of font (1/20 of a point)
     * @var integer
     */
    public $size = 10;

    /**
     * Bold style
     * @var integer
     */
    public $bold = self::FONT_NORMAL;

    /**
     * Bit specifiying if the font is italic.
     * @var integer
     */
    public $italic = 0;

    /**
     * Index to the cell's color
     * @var integer
     */
    public $color = 0x7FFF;

    /**
     * The text underline property
     * @var integer
     */
    public $underline = 0;

    /**
     * Bit specifiying if the font has strikeout.
     * @var integer
     */
    public $fontStrikeout = 0;

    /**
     * Bit specifiying if the font has outline.
     * @var integer
     */
    public $fontOutline = 0;

    /**
     * Bit specifiying if the font has shadow.
     * @var integer
     */
    public $fontShadow = 0;

    /**
     * 2 bytes specifiying the script type for the font.
     * @var integer
     */
    public $fontScript = 0;

    /**
     * Byte specifiying the font family.
     * @var integer
     */
    public $fontFamily = 0;

    /**
     * Byte specifiying the font charset.
     * @var integer
     */
    public $fontCharset = 0;

    /**
     * An index (2 bytes) to a FORMAT record (number format).
     * @var integer
     */
    public $numFormat = 0;

    /**
     * Bit specifying if formulas are hidden.
     * @var integer
     */
    public $hidden = 0;

    /**
     * Bit specifying if the cell is locked.
     * @var integer
     */
    public $locked = 0;

    /**
     * The three bits specifying the text horizontal alignment.
     * @var integer
     */
    public $textHorAlign = 0;

    /**
     * Bit specifying if the text is wrapped at the right border.
     * @var integer
     */
    public $textWrap = 0;

    /**
     * The three bits specifying the text vertical alignment.
     * @var integer
     */
    public $textVertAlign = 2;

    /**
     * 1 bit, apparently not used.
     * @var integer
     */
    public $textJustlast = 0;

    /**
     * The two bits specifying the text rotation.
     * @var integer
     */
    public $rotation = 0;

    /**
     * The cell's foreground color.
     * @var integer
     */
    public $fgColor = 0x40;

    /**
     * The cell's background color.
     * @var integer
     */
    public $bgColor = 0x41;

    /**
     * The cell's background fill pattern.
     * @var integer
     */
    public $pattern = 0;

    /**
     * Style of the bottom border of the cell
     * @var integer
     */
    public $bottom = 0;

    /**
     * Color of the bottom border of the cell.
     * @var integer
     */
    public $bottomColor = 0x40;

    /**
     * Style of the top border of the cell
     * @var integer
     */
    public $top = 0;

    /**
     * Color of the top border of the cell.
     * @var integer
     */
    public $topColor = 0x40;

    /**
     * Style of the left border of the cell
     * @var integer
     */
    public $left = 0;

    /**
     * Color of the left border of the cell.
     * @var integer
     */
    public $leftColor = 0x40;

    /**
     * Style of the right border of the cell
     * @var integer
     */
    public $right = 0;

    /**
     * Color of the right border of the cell.
     * @var integer
     */
    public $rightColor = 0x40;

    public $diag = 0;
    public $diagColor = 0x40;

    protected $horAlignMap = array(
        'left' => 1,
        'centre' => 2,
        'center' => 2,
        'right' => 3,
        'fill' => 4,
        'justify' => 5,
        'merge' => 6,
        'equal_space' => 7
    );

    protected $vertAlignMap = array(
        'top' => 0,
        'vcentre' => 1,
        'vcenter' => 1,
        'center' => 1,
        'bottom' => 2,
        'vjustify' => 3,
        'justify' => 3,
        'vequal_space' => 4,
        'equal_space' => 4
    );

    protected $rotationMap = array(
        0 => array(
            Biff5::VERSION => 0,
            Biff8::VERSION => 0,
        ),
        90 => array(
            Biff5::VERSION => 3,
            Biff8::VERSION => 180,
        ),
        270 => array(
            Biff5::VERSION => 2,
            Biff8::VERSION => 90,
        ),
        -1 => array(
            Biff5::VERSION => 1,
            Biff8::VERSION => 255,
        )
    );

    /**
     * @param integer $version
     * @param integer $byteOrder
     * @param integer $index the XF index for the format.
     * @param array $properties array with properties to be set on initialization.
     */
    public function __construct($version, $byteOrder, $index = 0, $properties = array())
    {
        $this->xfIndex = $index;
        $this->version = $version;
        $this->byteOrder = $byteOrder;

        $this->setProperties($properties);
    }

    /**
     * @param array $properties
     */
    protected function setProperties($properties)
    {
        foreach ($properties as $property => $value) {
            $methodName = 'set' . ucwords($property);
            if (method_exists($this, $methodName)) {
                $this->$methodName($value);
            }
        }
    }

    /**
     * Generate an Excel BIFF XF record (style or cell).
     *
     * @param string $style The type of the XF record ('style' or 'cell').
     * @return string The XF record data
     */
    public function getXf($style)
    {
        $record = new Record\Xf($this->version, $this->byteOrder);
        return $record->getData($this, $style);
    }

    /**
     * Generate an Excel BIFF FONT record.
     *
     * @return string The FONT record
     */
    public function getFontRecord()
    {
        $record = new Record\Font($this->version, $this->byteOrder);
        return $record->getData($this);
    }

    /**
     * Returns a unique hash key for a font.
     * The elements that form the key are arranged to increase the probability of
     * generating a unique key. Elements that hold a large range of numbers
     * (eg. color) are placed between two binary elements such as italic
     *
     * @return string A key for this font
     */
    public function getFontKey()
    {
        $key = "$this->fontName$this->size";
        $key .= "$this->fontScript$this->underline";
        $key .= "$this->fontStrikeout$this->bold$this->fontOutline";
        $key .= "$this->fontFamily$this->fontCharset";
        $key .= "$this->fontShadow$this->color$this->italic";
        $key = str_replace(' ', '_', $key);

        return $key;
    }

    /**
     * Returns the index used by Worksheet::xf()
     *
     * @return integer The index for the XF record
     */
    public function getXfIndex()
    {
        return $this->xfIndex;
    }

    /**
     * Used to convert a color
     * string into a number. Color range is 0..63 but we will restrict it
     * to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
     *
     * @param string|integer $name name of the color (i.e.: 'blue', 'red', etc..). Optional.
     *
     * @return integer The color index
     */
    protected function getColor($name = '')
    {
        return Palette::getColor($name);
    }

    /**
     * Set cell alignment.
     *
     * @param string $location alignment for the cell ('left', 'right', etc...).
     */
    public function setAlign($location)
    {
        $this->setHAlign($location);
        $this->setVAlign($location);
    }

    /**
     * Set cell horizontal alignment.
     *
     * @param string $location alignment for the cell ('left', 'right', etc...).
     */
    public function setHAlign($location)
    {
        $location = strtolower($location);
        if (isset($this->horAlignMap[$location])) {
            $this->textHorAlign = $this->horAlignMap[$location];
        }
    }

    /**
     * Set cell vertical alignment.
     *
     * @param string $location alignment for the cell ('top', 'vleft', 'vright', etc...).
     */
    public function setVAlign($location)
    {
        $location = strtolower($location);
        if (isset($this->vertAlignMap[$location])) {
            $this->textVertAlign = $this->vertAlignMap[$location];
        }
    }

    /**
     * Sets the boldness of the text.
     */
    public function setBold()
    {
        $this->bold = self::FONT_BOLD;
    }

    /**
     *
     */
    public function setNormal()
    {
        $this->bold = self::FONT_NORMAL;
    }

    /**
     * Sets the width for the bottom border of the cell
     *
     * @param integer $style style of the cell border (BORDER_THIN or BORDER_THICK).
     */
    public function setBottom($style)
    {
        $this->bottom = $style;
    }

    /**
     * Sets the width for the top border of the cell
     *
     * @param integer $style style of the cell top border (BORDER_THIN or BORDER_THICK).
     */
    public function setTop($style)
    {
        $this->top = $style;
    }

    /**
     * Sets the width for the left border of the cell
     *
     * @param integer $style style of the cell left border (BORDER_THIN or BORDER_THICK).
     */
    public function setLeft($style)
    {
        $this->left = $style;
    }

    /**
     * Sets the width for the right border of the cell
     *
     * @param integer $style style of the cell right border (BORDER_THIN or BORDER_THICK).
     */
    public function setRight($style)
    {
        $this->right = $style;
    }

    /**
     * Set cells borders to the same style
     *
     * @param integer $style style to apply for all cell borders (BORDER_THIN or BORDER_THICK).
     */
    public function setBorder($style)
    {
        $this->setBottom($style);
        $this->setTop($style);
        $this->setLeft($style);
        $this->setRight($style);
    }

    /**
     * Sets all the cell's borders to the same color
     *
     * @param string|integer $color The color we are setting. Either a string (like 'blue'),
     *                     or an integer (range is [8...63]).
     */
    public function setBorderColor($color)
    {
        $this->setBottomColor($color);
        $this->setTopColor($color);
        $this->setLeftColor($color);
        $this->setRightColor($color);
    }

    /**
     * Sets the cell's bottom border color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setBottomColor($color)
    {
        $this->bottomColor = $this->getColor($color);
    }

    /**
     * Sets the cell's top border color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setTopColor($color)
    {
        $this->topColor = $this->getColor($color);
    }

    /**
     * Sets the cell's left border color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setLeftColor($color)
    {
        $this->leftColor = $this->getColor($color);
    }

    /**
     * Sets the cell's right border color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setRightColor($color)
    {
        $this->rightColor = $this->getColor($color);
    }


    /**
     * Sets the cell's foreground color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setFgColor($color)
    {
        $this->fgColor = $this->getColor($color);
        if ($this->pattern == Fill::PATTERN_NONE) {
            $this->setPattern(Fill::PATTERN_SOLID);
        }
    }

    /**
     * Sets the cell's background color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setBgColor($color)
    {
        $this->bgColor = $this->getColor($color);
        if ($this->pattern == Fill::PATTERN_NONE) {
            $this->setPattern(Fill::PATTERN_SOLID);
        }
    }

    /**
     * Sets the cell's color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setColor($color)
    {
        $this->color = $this->getColor($color);
    }

    /**
     * Sets the fill pattern attribute of a cell
     *
     * @param integer $pattern Optional. Defaults to 1. Meaningful values are: 0-18,
     *                     0 meaning no background.
     */
    public function setPattern($pattern = Fill::PATTERN_SOLID)
    {
        $this->pattern = $pattern;
    }

    /**
     * Sets the underline of the text
     *
     * @param integer $underline The value for underline. Possible values are:
     * UNDERLINE_ONCE => underline, UNDERLINE_TWICE => double underline.
     */
    public function setUnderline($underline)
    {
        $this->underline = $underline;
    }

    /**
     * Sets the font style as italic
     *
     */
    public function setItalic()
    {
        $this->italic = 1;
    }

    /**
     * Sets the font size
     *
     * @param integer $size The font size (in pixels I think).
     */
    public function setSize($size)
    {
        $this->size = $size;
    }

    /**
     * Sets text wrapping
     *
     */
    public function setTextWrap()
    {
        $this->textWrap = 1;
    }

    /**
     * Sets the orientation of the text
     *
     * @param integer $angle The rotation angle for the text (clockwise). Possible
     * values are: 0, 90, 270 and -1 for stacking top-to-bottom.
     * @throws \Exception
     */
    public function setTextRotation($angle)
    {
        if (!isset($this->rotationMap[$angle])) {
            throw new \Exception(
                "Invalid value for angle." .
                " Possible values are: 0, 90, 270 and -1 " .
                "for stacking top-to-bottom."
            );
        }

        $this->rotation = $this->rotationMap[$angle][$this->version];
    }

    /**
     * Sets the numeric format.
     * It can be date, time, currency, etc...
     *
     * @param integer $numFormat The numeric format.
     */
    public function setNumFormat($numFormat)
    {
        $this->numFormat = $numFormat;
    }

    /**
     * Sets font as strikeout.
     *
     */
    public function setStrikeOut()
    {
        $this->fontStrikeout = 1;
    }

    /**
     * Sets outlining for a font.
     *
     */
    public function setOutLine()
    {
        $this->fontOutline = 1;
    }

    /**
     * Sets font as shadow.
     *
     */
    public function setShadow()
    {
        $this->fontShadow = 1;
    }

    /**
     * Sets the script type of the text
     *
     * @param integer $script The value for script type. Possible values are:
     * SCRIPT_SUPER => superscript, SCRIPT_SUB => subscript.
     */
    public function setScript($script)
    {
        $this->fontScript = $script;
    }

    /**
     * Locks a cell.
     */
    public function setLocked()
    {
        $this->locked = 1;
    }

    /**
     * Unlocks a cell. Useful for unprotecting particular cells of a protected sheet.
     */
    public function setUnLocked()
    {
        $this->locked = 0;
    }

    /**
     * Sets the font family name.
     *
     * @param string $fontFamily The font family name. Possible values are:
     *                           'Times New Roman', 'Arial', 'Courier'.
     */
    public function setFontFamily($fontFamily)
    {
        $this->fontName = $fontFamily;
    }

    /**
     * @return int
     */
    public function getVersion()
    {
        return $this->version;
    }

    /**
     * @return bool
     */
    public function isBuiltInNumFormat()
    {
        return preg_match("/^\d+$/", $this->numFormat) === 1;
    }

    /**
     * @return bool
     */
    public function isZeroStringNumFormat()
    {
        return preg_match("/^0+\d/", $this->numFormat) === 1;
    }
}
