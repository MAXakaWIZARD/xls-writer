<?php

namespace Xls;

class Format
{
    const BORDER_NONE = 0;
    const BORDER_THIN = 1;
    const BORDER_THICK = 2;

    /**
     * The index given by the workbook when creating a new format.
     * @var integer
     */
    public $xfIndex;

    /**
     * An index (2 bytes) to a FORMAT record (number format).
     * @var integer
     */
    protected $numFormat = NumberFormat::TYPE_GENERAL;

    /**
     * number format index
     * @var integer
     */
    protected $numFormatIndex;

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
    public $textVertAlign;

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
        0 => 0,
        90 => 180,
        270 => 90,
        -1 => 255
    );

    protected $borders = array(
        'top' => array(
            'style' => self::BORDER_NONE,
            'color' => 0
        ),
        'right' => array(
            'style' => self::BORDER_NONE,
            'color' => 0
        ),
        'bottom' => array(
            'style' => self::BORDER_NONE,
            'color' => 0
        ),
        'left' => array(
            'style' => self::BORDER_NONE,
            'color' => 0
        )
    );

    /**
     * @var Font
     */
    protected $font;

    /**
     * @param integer $index the XF index for the format.
     * @param array $properties array with properties to be set on initialization.
     */
    public function __construct($index = 0, $properties = array())
    {
        $this->xfIndex = $index;

        $this->font = new Font();

        $this->setVAlign('bottom');
        $this->setProperties($properties);
    }

    /**
     * @param array $properties
     */
    protected function setProperties($properties)
    {
        foreach ($properties as $property => $value) {
            $this->setProperty($property, $value);
        }
    }

    /**
     * @param $property
     * @param $value
     */
    protected function setProperty($property, $value)
    {
        $propertyParts = explode('.', $property);
        if (count($propertyParts) === 2
            && $propertyParts[0] === 'font'
        ) {
            $object = $this->getFont();
            $property = $propertyParts[1];
        } else {
            $object = $this;
        }

        $methodName = 'set' . ucwords($property);
        if (method_exists($object, $methodName)) {
            $object->$methodName($value);
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
        $record = new Record\Xf();
        return $record->getData($this, $style);
    }

    /**
     * @return Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Generate an Excel BIFF FONT record.
     *
     * @return string The FONT record
     */
    public function getFontRecord()
    {
        $record = new Record\Font();

        return $record->getData($this->getFont());
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
     * Sets the style for the bottom border of the cell
     *
     * @param integer $style style of the cell border (BORDER_THIN or BORDER_THICK).
     * @param string|integer $color The color we are setting. Either a string (like 'blue'),
     *                     or an integer (range is [8...63]).
     */
    public function setBorderBottom($style, $color = 0x40)
    {
        $this->setBorderInternal('bottom', $style, $color);
    }

    /**
     * Sets the style for the top border of the cell
     *
     * @param integer $style style of the cell top border (BORDER_THIN or BORDER_THICK).
     * @param string|integer $color The color we are setting. Either a string (like 'blue'),
     *                     or an integer (range is [8...63]).
     */
    public function setBorderTop($style, $color = 0x40)
    {
        $this->setBorderInternal('top', $style, $color);
    }

    /**
     * Sets the style for the left border of the cell
     *
     * @param integer $style style of the cell left border (BORDER_THIN or BORDER_THICK).
     * @param string|integer $color The color we are setting. Either a string (like 'blue'),
     *                     or an integer (range is [8...63]).
     */
    public function setBorderLeft($style, $color = 0x40)
    {
        $this->setBorderInternal('left', $style, $color);
    }

    /**
     * Sets the style for the right border of the cell
     *
     * @param integer $style style of the cell right border (BORDER_THIN or BORDER_THICK).
     * @param string|integer $color The color we are setting. Either a string (like 'blue'),
     *                     or an integer (range is [8...63]).
     */
    public function setBorderRight($style, $color = 0x40)
    {
        $this->setBorderInternal('right', $style, $color);
    }

    /**
     * Set cells borders to the same style
     *
     * @param integer $style style to apply for all cell borders (BORDER_THIN or BORDER_THICK).
     * @param string|integer $color The color we are setting. Either a string (like 'blue'),
     *                     or an integer (range is [8...63]).
     */
    public function setBorder($style, $color = 0x40)
    {
        $this->setBorderBottom($style, $color);
        $this->setBorderTop($style, $color);
        $this->setBorderLeft($style, $color);
        $this->setBorderRight($style, $color);
    }

    /**
     * Sets the style for the bottom border of the cell
     * @param string $side
     * @param integer $style style of the cell border (BORDER_THIN or BORDER_THICK).
     * @param string|integer $color The color we are setting. Either a string (like 'blue'),
     *                     or an integer (range is [8...63]).
     */
    protected function setBorderInternal($side, $style, $color = 0x40)
    {
        $this->borders[$side]['style'] = $style;

        if (!is_null($color)) {
            $this->borders[$side]['color'] = Palette::getColor($color);
        }
    }

    /**
     * @param $side
     *
     * @return integer|null
     */
    public function getBorderStyle($side)
    {
        return (isset($this->borders[$side])) ? $this->borders[$side]['style'] : null;
    }

    /**
     * @param $side
     *
     * @return integer|null
     */
    public function getBorderColor($side)
    {
        return (isset($this->borders[$side])) ? $this->borders[$side]['color'] : null;
    }

    /**
     * Sets the cell's foreground color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setFgColor($color)
    {
        $this->fgColor = Palette::getColor($color);
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
        $this->bgColor = Palette::getColor($color);
        if ($this->pattern == Fill::PATTERN_NONE) {
            $this->setPattern(Fill::PATTERN_SOLID);
        }
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

        $this->rotation = $this->rotationMap[$angle];
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
     * @return int
     */
    public function getNumFormat()
    {
        return $this->numFormat;
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
     * @return int
     */
    public function getNumFormatIndex()
    {
        return $this->numFormatIndex;
    }

    /**
     * @param int $numFormatIndex
     */
    public function setNumFormatIndex($numFormatIndex)
    {
        $this->numFormatIndex = $numFormatIndex;
    }
}
