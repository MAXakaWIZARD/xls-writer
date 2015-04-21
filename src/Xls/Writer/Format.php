<?php
/*
*  Module written/ported by Xavier Noguer <xnoguer@rezebra.com>
*
*  The majority of this is _NOT_ my code.  I simply ported it from the
*  PERL Spreadsheet::WriteExcel module.
*
*  The author of the Spreadsheet::WriteExcel module is John McNamara
*  <jmcnamara@cpan.org>
*
*  I _DO_ maintain this code, and John McNamara has nothing to do with the
*  porting of this code to PHP.  Any questions directly related to this
*  class library should be directed to me.
*
*  License Information:
*
*    Spreadsheet_Excel_Writer:  A library for generating Excel Spreadsheets
*    Copyright (c) 2002-2003 Xavier Noguer xnoguer@rezebra.com
*
*    This library is free software; you can redistribute it and/or
*    modify it under the terms of the GNU Lesser General Public
*    License as published by the Free Software Foundation; either
*    version 2.1 of the License, or (at your option) any later version.
*
*    This library is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
*    Lesser General Public License for more details.
*
*    You should have received a copy of the GNU Lesser General Public
*    License along with this library; if not, write to the Free Software
*    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*/

namespace Xls\Writer;

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

    /**
     * The index given by the workbook when creating a new format.
     * @var integer
     */
    public $xf_index;

    /**
     * Index to the FONT record.
     * @var integer
     */
    public $font_index;

    /**
     * The font name (ASCII).
     * @var string
     */
    public $font_name;

    /**
     * Height of font (1/20 of a point)
     * @var integer
     */
    public $size;

    /**
     * Bold style
     * @var integer
     */
    public $bold;

    /**
     * Bit specifiying if the font is italic.
     * @var integer
     */
    public $italic;

    /**
     * Index to the cell's color
     * @var integer
     */
    public $color;

    /**
     * The text underline property
     * @var integer
     */
    public $underline;

    /**
     * Bit specifiying if the font has strikeout.
     * @var integer
     */
    public $font_strikeout;

    /**
     * Bit specifiying if the font has outline.
     * @var integer
     */
    public $font_outline;

    /**
     * Bit specifiying if the font has shadow.
     * @var integer
     */
    public $font_shadow;

    /**
     * 2 bytes specifiying the script type for the font.
     * @var integer
     */
    public $font_script;

    /**
     * Byte specifiying the font family.
     * @var integer
     */
    public $font_family;

    /**
     * Byte specifiying the font charset.
     * @var integer
     */
    public $font_charset;

    /**
     * An index (2 bytes) to a FORMAT record (number format).
     * @var integer
     */
    public $num_format;

    /**
     * Bit specifying if formulas are hidden.
     * @var integer
     */
    public $hidden;

    /**
     * Bit specifying if the cell is locked.
     * @var integer
     */
    public $locked;

    /**
     * The three bits specifying the text horizontal alignment.
     * @var integer
     */
    public $text_h_align;

    /**
     * Bit specifying if the text is wrapped at the right border.
     * @var integer
     */
    public $text_wrap;

    /**
     * The three bits specifying the text vertical alignment.
     * @var integer
     */
    public $text_v_align;

    /**
     * 1 bit, apparently not used.
     * @var integer
     */
    public $text_justlast;

    /**
     * The two bits specifying the text rotation.
     * @var integer
     */
    public $rotation;

    /**
     * The cell's foreground color.
     * @var integer
     */
    public $fg_color;

    /**
     * The cell's background color.
     * @var integer
     */
    public $bg_color;

    /**
     * The cell's background fill pattern.
     * @var integer
     */
    public $pattern;

    /**
     * Style of the bottom border of the cell
     * @var integer
     */
    public $bottom;

    /**
     * Color of the bottom border of the cell.
     * @var integer
     */
    public $bottom_color;

    /**
     * Style of the top border of the cell
     * @var integer
     */
    public $top;

    /**
     * Color of the top border of the cell.
     * @var integer
     */
    public $top_color;

    /**
     * Style of the left border of the cell
     * @var integer
     */
    public $left;

    /**
     * Color of the left border of the cell.
     * @var integer
     */
    public $left_color;

    /**
     * Style of the right border of the cell
     * @var integer
     */
    public $right;

    /**
     * Color of the right border of the cell.
     * @var integer
     */
    public $right_color;

    /**
     * Constructor
     *
     * @param integer $index the XF index for the format.
     * @param array $properties array with properties to be set on initialization.
     */
    public function __construct($biffVersion, $index = 0, $properties = array())
    {
        $this->xf_index = $index;
        $this->biffVersion = $biffVersion;
        $this->font_index = 0;
        $this->font_name = 'Arial';
        $this->size = 10;
        $this->bold = 0x0190;
        $this->italic = 0;
        $this->color = 0x7FFF;
        $this->underline = 0;
        $this->font_strikeout = 0;
        $this->font_outline = 0;
        $this->font_shadow = 0;
        $this->font_script = 0;
        $this->font_family = 0;
        $this->font_charset = 0;

        $this->num_format = 0;

        $this->hidden = 0;
        $this->locked = 0;

        $this->text_h_align = 0;
        $this->text_wrap = 0;
        $this->text_v_align = 2;
        $this->text_justlast = 0;
        $this->rotation = 0;

        $this->fg_color = 0x40;
        $this->bg_color = 0x41;

        $this->pattern = 0;

        $this->bottom = 0;
        $this->top = 0;
        $this->left = 0;
        $this->right = 0;
        $this->diag = 0;

        $this->bottom_color = 0x40;
        $this->top_color = 0x40;
        $this->left_color = 0x40;
        $this->right_color = 0x40;
        $this->diag_color = 0x40;

        // Set properties passed to Workbook::addFormat()
        foreach ($properties as $property => $value) {
            if (method_exists($this, 'set' . ucwords($property))) {
                $method_name = 'set' . ucwords($property);
                $this->$method_name($value);
            }
        }
    }


    /**
     * Generate an Excel BIFF XF record (style or cell).
     *
     * @param string $style The type of the XF record ('style' or 'cell').
     * @return string The XF record
     */
    public function getXf($style)
    {
        // Set the type of the XF record and some of the attributes.
        if ($style == 'style') {
            $style = 0xFFF5;
        } else {
            $style = $this->locked;
            $style |= $this->hidden << 1;
        }

        // Flags to indicate if attributes have been set.
        $atr_num = ($this->num_format != 0) ? 1 : 0;
        $atr_fnt = ($this->font_index != 0) ? 1 : 0;
        $atr_alc = ($this->text_wrap) ? 1 : 0;
        $atr_bdr = ($this->bottom
            || $this->top
            || $this->left
            || $this->right) ? 1 : 0;
        $atr_pat = (($this->fg_color != 0x40)
            || ($this->bg_color != 0x41)
            || $this->pattern) ? 1 : 0;
        $atr_prot = $this->locked | $this->hidden;

        // Zero the default border colour if the border has not been set.
        if ($this->bottom == 0) {
            $this->bottom_color = 0;
        }
        if ($this->top == 0) {
            $this->top_color = 0;
        }
        if ($this->right == 0) {
            $this->right_color = 0;
        }
        if ($this->left == 0) {
            $this->left_color = 0;
        }
        if ($this->diag == 0) {
            $this->diag_color = 0;
        }

        $record = 0x00E0; // Record identifier
        if ($this->biffVersion === BIFFwriter::VERSION_5) {
            $length = 0x0010; // Number of bytes to follow
        } else {
            $length = 0x0014;
        }

        $ifnt = $this->font_index; // Index to FONT record
        $ifmt = $this->num_format; // Index to FORMAT record
        if ($this->biffVersion === BIFFwriter::VERSION_5) {
            $align = $this->text_h_align; // Alignment
            $align |= $this->text_wrap << 3;
            $align |= $this->text_v_align << 4;
            $align |= $this->text_justlast << 7;
            $align |= $this->rotation << 8;
            $align |= $atr_num << 10;
            $align |= $atr_fnt << 11;
            $align |= $atr_alc << 12;
            $align |= $atr_bdr << 13;
            $align |= $atr_pat << 14;
            $align |= $atr_prot << 15;

            $icv = $this->fg_color; // fg and bg pattern colors
            $icv |= $this->bg_color << 7;

            $fill = $this->pattern; // Fill and border line style
            $fill |= $this->bottom << 6;
            $fill |= $this->bottom_color << 9;

            $border1 = $this->top; // Border line style and color
            $border1 |= $this->left << 3;
            $border1 |= $this->right << 6;
            $border1 |= $this->top_color << 9;

            $border2 = $this->left_color; // Border color
            $border2 |= $this->right_color << 7;

            $header = pack("vv", $record, $length);
            $data = pack(
                "vvvvvvvv",
                $ifnt,
                $ifmt,
                $style,
                $align,
                $icv,
                $fill,
                $border1,
                $border2
            );
        } else {
            $align = $this->text_h_align; // Alignment
            $align |= $this->text_wrap << 3;
            $align |= $this->text_v_align << 4;
            $align |= $this->text_justlast << 7;

            $used_attrib = $atr_num << 2;
            $used_attrib |= $atr_fnt << 3;
            $used_attrib |= $atr_alc << 4;
            $used_attrib |= $atr_bdr << 5;
            $used_attrib |= $atr_pat << 6;
            $used_attrib |= $atr_prot << 7;

            $icv = $this->fg_color; // fg and bg pattern colors
            $icv |= $this->bg_color << 7;

            $border1 = $this->left; // Border line style and color
            $border1 |= $this->right << 4;
            $border1 |= $this->top << 8;
            $border1 |= $this->bottom << 12;
            $border1 |= $this->left_color << 16;
            $border1 |= $this->right_color << 23;
            $diag_tl_to_rb = 0;
            $diag_tr_to_lb = 0;
            $border1 |= $diag_tl_to_rb << 30;
            $border1 |= $diag_tr_to_lb << 31;

            $border2 = $this->top_color; // Border color
            $border2 |= $this->bottom_color << 7;
            $border2 |= $this->diag_color << 14;
            $border2 |= $this->diag << 21;
            $border2 |= $this->pattern << 26;

            $header = pack("vv", $record, $length);

            $rotation = $this->rotation;
            $biff8_options = 0x00;
            $data = pack("vvvC", $ifnt, $ifmt, $style, $align);
            $data .= pack("CCC", $rotation, $biff8_options, $used_attrib);
            $data .= pack("VVv", $border1, $border2, $icv);
        }

        return ($header . $data);
    }

    /**
     * Generate an Excel BIFF FONT record.
     *
     * @return string The FONT record
     */
    public function getFont()
    {
        $dyHeight = $this->size * 20; // Height of font (1/20 of a point)
        $icv = $this->color; // Index to color palette
        $bls = $this->bold; // Bold style
        $sss = $this->font_script; // Superscript/subscript
        $uls = $this->underline; // Underline
        $bFamily = $this->font_family; // Font family
        $bCharSet = $this->font_charset; // Character set
        $encoding = 0;

        $cch = strlen($this->font_name); // Length of font name
        $record = 0x31; // Record identifier
        if ($this->biffVersion === BIFFwriter::VERSION_5) {
            $length = 0x0F + $cch; // Record length
        } else {
            $length = 0x10 + $cch;
        }

        $reserved = 0x00; // Reserved
        $grbit = 0x00; // Font attributes
        if ($this->italic) {
            $grbit |= 0x02;
        }
        if ($this->font_strikeout) {
            $grbit |= 0x08;
        }
        if ($this->font_outline) {
            $grbit |= 0x10;
        }
        if ($this->font_shadow) {
            $grbit |= 0x20;
        }

        $header = pack("vv", $record, $length);
        if ($this->biffVersion === BIFFwriter::VERSION_5) {
            $data = pack(
                "vvvvvCCCCC",
                $dyHeight,
                $grbit,
                $icv,
                $bls,
                $sss,
                $uls,
                $bFamily,
                $bCharSet,
                $reserved,
                $cch
            );
        } else {
            $data = pack(
                "vvvvvCCCCCC",
                $dyHeight,
                $grbit,
                $icv,
                $bls,
                $sss,
                $uls,
                $bFamily,
                $bCharSet,
                $reserved,
                $cch,
                $encoding
            );
        }

        return ($header . $data . $this->font_name);
    }

    /**
     * Returns a unique hash key for a font.
     * Used by Workbook::storeAllFonts()
     *
     * The elements that form the key are arranged to increase the probability of
     * generating a unique key. Elements that hold a large range of numbers
     * (eg. _color) are placed between two binary elements such as _italic
     *
     * @return string A key for this font
     */
    public function getFontKey()
    {
        $key = "$this->font_name$this->size";
        $key .= "$this->font_script$this->underline";
        $key .= "$this->font_strikeout$this->bold$this->font_outline";
        $key .= "$this->font_family$this->font_charset";
        $key .= "$this->font_shadow$this->color$this->italic";
        $key = str_replace(' ', '_', $key);

        return ($key);
    }

    /**
     * Returns the index used by Worksheet::xf()
     *
     * @return integer The index for the XF record
     */
    public function getXfIndex()
    {
        return $this->xf_index;
    }

    /**
     * Used in conjunction with the set_xxx_color methods to convert a color
     * string into a number. Color range is 0..63 but we will restrict it
     * to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
     *
     * @param string $name_color name of the color (i.e.: 'blue', 'red', etc..). Optional.
     * @return integer The color index
     */
    protected function getColor($name_color = '')
    {
        $colors = array(
            'aqua' => 0x07,
            'cyan' => 0x07,
            'black' => 0x00,
            'blue' => 0x04,
            'brown' => 0x10,
            'magenta' => 0x06,
            'fuchsia' => 0x06,
            'gray' => 0x17,
            'grey' => 0x17,
            'green' => 0x11,
            'lime' => 0x03,
            'navy' => 0x12,
            'orange' => 0x35,
            'purple' => 0x14,
            'red' => 0x02,
            'silver' => 0x16,
            'white' => 0x01,
            'yellow' => 0x05
        );

        // Return the default color, 0x7FFF, if undef,
        if ($name_color === '') {
            return (0x7FFF);
        }

        // or the color string converted to an integer,
        if (isset($colors[$name_color])) {
            return ($colors[$name_color]);
        }

        // or the default color if string is unrecognised,
        if (preg_match("/\D/", $name_color)) {
            return (0x7FFF);
        }

        // or the default color if arg is outside range,
        if ($name_color > 63) {
            return (0x7FFF);
        }

        // or an integer in the valid range
        return $name_color;
    }

    /**
     * Set cell alignment.
     *
     * @param string $location alignment for the cell ('left', 'right', etc...).
     */
    public function setAlign($location)
    {
        if (preg_match("/\d/", $location)) {
            return; // Ignore numbers
        }

        $location = strtolower($location);

        if ($location == 'left') {
            $this->text_h_align = 1;
        }
        if ($location == 'centre') {
            $this->text_h_align = 2;
        }
        if ($location == 'center') {
            $this->text_h_align = 2;
        }
        if ($location == 'right') {
            $this->text_h_align = 3;
        }
        if ($location == 'fill') {
            $this->text_h_align = 4;
        }
        if ($location == 'justify') {
            $this->text_h_align = 5;
        }
        if ($location == 'merge') {
            $this->text_h_align = 6;
        }
        if ($location == 'equal_space') { // For T.K.
            $this->text_h_align = 7;
        }
        if ($location == 'top') {
            $this->text_v_align = 0;
        }
        if ($location == 'vcentre') {
            $this->text_v_align = 1;
        }
        if ($location == 'vcenter') {
            $this->text_v_align = 1;
        }
        if ($location == 'bottom') {
            $this->text_v_align = 2;
        }
        if ($location == 'vjustify') {
            $this->text_v_align = 3;
        }
        if ($location == 'vequal_space') { // For T.K.
            $this->text_v_align = 4;
        }
    }

    /**
     * Set cell horizontal alignment.
     *
     * @param string $location alignment for the cell ('left', 'right', etc...).
     */
    public function setHAlign($location)
    {
        if (preg_match("/\d/", $location)) {
            return; // Ignore numbers
        }

        $location = strtolower($location);

        if ($location == 'left') {
            $this->text_h_align = 1;
        }
        if ($location == 'centre') {
            $this->text_h_align = 2;
        }
        if ($location == 'center') {
            $this->text_h_align = 2;
        }
        if ($location == 'right') {
            $this->text_h_align = 3;
        }
        if ($location == 'fill') {
            $this->text_h_align = 4;
        }
        if ($location == 'justify') {
            $this->text_h_align = 5;
        }
        if ($location == 'merge') {
            $this->text_h_align = 6;
        }
        if ($location == 'equal_space') { // For T.K.
            $this->text_h_align = 7;
        }
    }

    /**
     * Set cell vertical alignment.
     *
     * @param string $location alignment for the cell ('top', 'vleft', 'vright', etc...).
     */
    public function setVAlign($location)
    {
        if (preg_match("/\d/", $location)) {
            return; // Ignore numbers
        }

        $location = strtolower($location);

        if ($location == 'top') {
            $this->text_v_align = 0;
        }
        if ($location == 'vcentre') {
            $this->text_v_align = 1;
        }
        if ($location == 'vcenter') {
            $this->text_v_align = 1;
        }
        if ($location == 'bottom') {
            $this->text_v_align = 2;
        }
        if ($location == 'vjustify') {
            $this->text_v_align = 3;
        }
        if ($location == 'vequal_space') { // For T.K.
            $this->text_v_align = 4;
        }
    }

    /**
     * This is an alias for the unintuitive setAlign('merge')
     *
     */
    public function setMerge()
    {
        $this->setAlign('merge');
    }

    /**
     * Sets the boldness of the text.
     * Bold has a range 100..1000.
     * 0 (400) is normal. 1 (700) is bold.
     *
     * @param integer $weight Weight for the text, 0 maps to 400 (normal text),
    1 maps to 700 (bold text). Valid range is: 100-1000.
    It's Optional, default is 1 (bold).
     */
    public function setBold($weight = 1)
    {
        if ($weight == 1) {
            $weight = 0x2BC; // Bold text
        }
        if ($weight == 0) {
            $weight = 0x190; // Normal text
        }
        if ($weight < 0x064) {
            $weight = 0x190; // Lower bound
        }
        if ($weight > 0x3E8) {
            $weight = 0x190; // Upper bound
        }
        $this->bold = $weight;
    }


    /************************************
     * FUNCTIONS FOR SETTING CELLS BORDERS
     */

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


    /*******************************************
     * FUNCTIONS FOR SETTING CELLS BORDERS COLORS
     */

    /**
     * Sets all the cell's borders to the same color
     *
     * @param mixed $color The color we are setting. Either a string (like 'blue'),
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
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setBottomColor($color)
    {
        $value = $this->getColor($color);
        $this->bottom_color = $value;
    }

    /**
     * Sets the cell's top border color
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setTopColor($color)
    {
        $value = $this->getColor($color);
        $this->top_color = $value;
    }

    /**
     * Sets the cell's left border color
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setLeftColor($color)
    {
        $value = $this->getColor($color);
        $this->left_color = $value;
    }

    /**
     * Sets the cell's right border color
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setRightColor($color)
    {
        $value = $this->getColor($color);
        $this->right_color = $value;
    }


    /**
     * Sets the cell's foreground color
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setFgColor($color)
    {
        $value = $this->getColor($color);
        $this->fg_color = $value;
        if ($this->pattern == 0) { // force color to be seen
            $this->pattern = 1;
        }
    }

    /**
     * Sets the cell's background color
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setBgColor($color)
    {
        $value = $this->getColor($color);
        $this->bg_color = $value;
        if ($this->pattern == 0) { // force color to be seen
            $this->pattern = 1;
        }
    }

    /**
     * Sets the cell's color
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setColor($color)
    {
        $value = $this->getColor($color);
        $this->color = $value;
    }

    /**
     * Sets the fill pattern attribute of a cell
     *
     * @param integer $arg Optional. Defaults to 1. Meaningful values are: 0-18,
     *                     0 meaning no background.
     */
    public function setPattern($arg = 1)
    {
        $this->pattern = $arg;
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
        $this->text_wrap = 1;
    }

    /**
     * Sets the orientation of the text
     *
     * @param integer $angle The rotation angle for the text (clockwise). Possible
    values are: 0, 90, 270 and -1 for stacking top-to-bottom.
     * @throws \Exception
     */
    public function setTextRotation($angle)
    {
        switch ($angle) {
            case 0:
                $this->rotation = 0;
                break;
            case 90:
                if ($this->biffVersion === BIFFwriter::VERSION_5) {
                    $this->rotation = 3;
                } else {
                    $this->rotation = 180;
                }
                break;
            case 270:
                if ($this->biffVersion == BIFFwriter::VERSION_5) {
                    $this->rotation = 2;
                } else {
                    $this->rotation = 90;
                }
                break;
            case -1:
                if ($this->biffVersion == BIFFwriter::VERSION_5) {
                    $this->rotation = 1;
                } else {
                    $this->rotation = 255;
                }
                break;
            default:
                throw new \Exception(
                    "Invalid value for angle." .
                    " Possible values are: 0, 90, 270 and -1 " .
                    "for stacking top-to-bottom."
                );
                break;
        }
    }

    /**
     * Sets the numeric format.
     * It can be date, time, currency, etc...
     *
     * @param integer $num_format The numeric format.
     */
    public function setNumFormat($num_format)
    {
        $this->num_format = $num_format;
    }

    /**
     * Sets font as strikeout.
     *
     */
    public function setStrikeOut()
    {
        $this->font_strikeout = 1;
    }

    /**
     * Sets outlining for a font.
     *
     */
    public function setOutLine()
    {
        $this->font_outline = 1;
    }

    /**
     * Sets font as shadow.
     *
     */
    public function setShadow()
    {
        $this->font_shadow = 1;
    }

    /**
     * Sets the script type of the text
     *
     * @param integer $script The value for script type. Possible values are:
     * SCRIPT_SUPER => superscript, SCRIPT_SUB => subscript.
     */
    public function setScript($script)
    {
        $this->font_script = $script;
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
     * @param string $fontfamily The font family name. Possible values are:
     *                           'Times New Roman', 'Arial', 'Courier'.
     */
    public function setFontFamily($fontFamily)
    {
        $this->font_name = $fontFamily;
    }
}
