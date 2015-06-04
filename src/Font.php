<?php

namespace Xls;

class Font
{
    const FONT_NORMAL = 400;
    const FONT_BOLD = 700;

    const UNDERLINE_NONE = 0;
    const UNDERLINE_ONCE = 1;
    const UNDERLINE_TWICE = 2;

    const SCRIPT_NONE = 0;
    const SCRIPT_SUPER = 1;
    const SCRIPT_SUB = 2;

    /**
     * Index to the FONT record.
     * @var integer
     */
    public $index = 0;

    /**
     * The font name (ASCII).
     * @var string
     */
    public $name = 'Arial';

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
    public $underline = self::UNDERLINE_NONE;

    /**
     * Bit specifiying if the font has strikeout.
     * @var integer
     */
    public $strikeout = 0;

    /**
     * Bit specifiying if the font has outline.
     * @var integer
     */
    public $outline = 0;

    /**
     * Bit specifiying if the font has shadow.
     * @var integer
     */
    public $shadow = 0;

    /**
     * 2 bytes specifiying the script type for the font.
     * @var integer
     */
    public $script = self::SCRIPT_NONE;

    /**
     * Sets the boldness of the text.
     * @param bool $bold
     *
     * @return Font
     */
    public function setBold($bold = true)
    {
        $this->bold = ($bold) ? self::FONT_BOLD : self::FONT_NORMAL;

        return $this;
    }

    /**
     * Sets the cell's font color
     *
     * @param string|integer $color either a string (like 'blue'), or an integer (range is [8...63]).
     *
     * @return Font
     */
    public function setColor($color)
    {
        $this->color = Palette::getColor($color);

        return $this;
    }

    /**
     * Sets the underline of the text
     *
     * @param integer $underline The value for underline. Possible values are:
     * UNDERLINE_ONCE => underline, UNDERLINE_TWICE => double underline.
     *
     * @return Font
     */
    public function setUnderline($underline)
    {
        $this->underline = $underline;

        return $this;
    }

    /**
     * Sets the font style as italic
     * @param bool $italic
     *
     * @return Font
     */
    public function setItalic($italic = true)
    {
        $this->italic = ($italic) ? 1 : 0;

        return $this;
    }

    /**
     * Sets the font size
     *
     * @param integer $size The font size (in pixels I think).
     *
     * @return Font
     */
    public function setSize($size)
    {
        $this->size = $size;

        return $this;
    }

    /**
     * Sets font as strikeout
     * @param bool $strikeout
     *
     * @return Font
     */
    public function setStrikeOut($strikeout = true)
    {
        $this->strikeout = ($strikeout) ? 1 : 0;

        return $this;
    }

    /**
     * Sets outlining for a font.
     * @param bool $outline
     *
     * @return Font
     */
    public function setOutLine($outline = true)
    {
        $this->outline = ($outline) ? 1 : 0;

        return $this;
    }

    /**
     * Sets font as shadow.
     * @param bool $shadow
     *
     * @return Font
     */
    public function setShadow($shadow = true)
    {
        $this->shadow = ($shadow) ? 1 : 0;

        return $this;
    }

    /**
     * @param bool $enable
     *
     * @return Font
     */
    public function setSuperScript($enable = true)
    {
        $this->script = ($enable) ? self::SCRIPT_SUPER : self::SCRIPT_NONE;

        return $this;
    }

    /**
     * @param bool $enable
     *
     * @return Font
     */
    public function setSubScript($enable = true)
    {
        $this->script = ($enable) ? self::SCRIPT_SUB : self::SCRIPT_NONE;

        return $this;
    }

    /**
     * Sets the font name.
     *
     * @param string $name The font name. Possible values are:
     *                           'Times New Roman', 'Arial', 'Courier'.
     *
     * @return Font
     */
    public function setName($name)
    {
        $this->name = $name;

        return $this;
    }

    /**
     * Returns a unique hash key for a font.
     * The elements that form the key are arranged to increase the probability of
     * generating a unique key. Elements that hold a large range of numbers
     * (eg. color) are placed between two binary elements such as italic
     *
     * @return string A key for this font
     */
    public function getKey()
    {
        $key = "$this->name$this->size";
        $key .= "$this->script$this->underline";
        $key .= "$this->strikeout$this->bold$this->outline";
        $key .= "$this->shadow$this->color$this->italic";
        $key = str_replace(' ', '_', $key);

        return $key;
    }

    /**
     * @return int
     */
    public function getIndex()
    {
        return $this->index;
    }

    /**
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * @return int
     */
    public function getSize()
    {
        return $this->size;
    }

    /**
     * @return int
     */
    public function getBold()
    {
        return $this->bold;
    }

    /**
     * @return int
     */
    public function getItalic()
    {
        return $this->italic;
    }

    /**
     * @return int
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * @return int
     */
    public function getUnderline()
    {
        return $this->underline;
    }

    /**
     * @return int
     */
    public function getStrikeout()
    {
        return $this->strikeout;
    }

    /**
     * @return int
     */
    public function getOutline()
    {
        return $this->outline;
    }

    /**
     * @return int
     */
    public function getShadow()
    {
        return $this->shadow;
    }

    /**
     * @return int
     */
    public function getScript()
    {
        return $this->script;
    }

    /**
     * @param int $index
     */
    public function setIndex($index)
    {
        $this->index = $index;
    }
}
