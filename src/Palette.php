<?php

namespace Xls;

class Palette
{
    protected static $colorsMap = array(
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

    protected static $x97Palette = array(
        array(0x00, 0x00, 0x00, 0x00), // 8
        array(0xff, 0xff, 0xff, 0x00), // 9
        array(0xff, 0x00, 0x00, 0x00), // 10
        array(0x00, 0xff, 0x00, 0x00), // 11
        array(0x00, 0x00, 0xff, 0x00), // 12
        array(0xff, 0xff, 0x00, 0x00), // 13
        array(0xff, 0x00, 0xff, 0x00), // 14
        array(0x00, 0xff, 0xff, 0x00), // 15
        array(0x80, 0x00, 0x00, 0x00), // 16
        array(0x00, 0x80, 0x00, 0x00), // 17
        array(0x00, 0x00, 0x80, 0x00), // 18
        array(0x80, 0x80, 0x00, 0x00), // 19
        array(0x80, 0x00, 0x80, 0x00), // 20
        array(0x00, 0x80, 0x80, 0x00), // 21
        array(0xc0, 0xc0, 0xc0, 0x00), // 22
        array(0x80, 0x80, 0x80, 0x00), // 23
        array(0x99, 0x99, 0xff, 0x00), // 24
        array(0x99, 0x33, 0x66, 0x00), // 25
        array(0xff, 0xff, 0xcc, 0x00), // 26
        array(0xcc, 0xff, 0xff, 0x00), // 27
        array(0x66, 0x00, 0x66, 0x00), // 28
        array(0xff, 0x80, 0x80, 0x00), // 29
        array(0x00, 0x66, 0xcc, 0x00), // 30
        array(0xcc, 0xcc, 0xff, 0x00), // 31
        array(0x00, 0x00, 0x80, 0x00), // 32
        array(0xff, 0x00, 0xff, 0x00), // 33
        array(0xff, 0xff, 0x00, 0x00), // 34
        array(0x00, 0xff, 0xff, 0x00), // 35
        array(0x80, 0x00, 0x80, 0x00), // 36
        array(0x80, 0x00, 0x00, 0x00), // 37
        array(0x00, 0x80, 0x80, 0x00), // 38
        array(0x00, 0x00, 0xff, 0x00), // 39
        array(0x00, 0xcc, 0xff, 0x00), // 40
        array(0xcc, 0xff, 0xff, 0x00), // 41
        array(0xcc, 0xff, 0xcc, 0x00), // 42
        array(0xff, 0xff, 0x99, 0x00), // 43
        array(0x99, 0xcc, 0xff, 0x00), // 44
        array(0xff, 0x99, 0xcc, 0x00), // 45
        array(0xcc, 0x99, 0xff, 0x00), // 46
        array(0xff, 0xcc, 0x99, 0x00), // 47
        array(0x33, 0x66, 0xff, 0x00), // 48
        array(0x33, 0xcc, 0xcc, 0x00), // 49
        array(0x99, 0xcc, 0x00, 0x00), // 50
        array(0xff, 0xcc, 0x00, 0x00), // 51
        array(0xff, 0x99, 0x00, 0x00), // 52
        array(0xff, 0x66, 0x00, 0x00), // 53
        array(0x66, 0x66, 0x99, 0x00), // 54
        array(0x96, 0x96, 0x96, 0x00), // 55
        array(0x00, 0x33, 0x66, 0x00), // 56
        array(0x33, 0x99, 0x66, 0x00), // 57
        array(0x00, 0x33, 0x00, 0x00), // 58
        array(0x33, 0x33, 0x00, 0x00), // 59
        array(0x99, 0x33, 0x00, 0x00), // 60
        array(0x99, 0x33, 0x66, 0x00), // 61
        array(0x33, 0x33, 0x99, 0x00), // 62
        array(0x33, 0x33, 0x33, 0x00), // 63
    );

    /**
     * Return Excel 97+ default palette
     * @return array
     */
    public static function getXl97Palette()
    {
        return self::$x97Palette;
    }

    /**
     * @param $index
     * @param $red
     * @param $green
     * @param $blue
     *
     * @throws \Exception
     */
    public static function validateColor($index, $red, $green, $blue)
    {
        // Check that the colour index is the right range
        if ($index < 8 || $index > 64) {
            throw new \Exception("Color index $index outside range: 8 <= index <= 64");
        }

        // Check that the colour components are in the right range
        if (($red < 0 || $red > 255)
            || ($green < 0 || $green > 255)
            || ($blue < 0 || $blue > 255)
        ) {
            throw new \Exception("Color component outside range: 0 <= color <= 255");
        }
    }

    /**
     * @param $name
     *
     * @return bool
     */
    public static function isValidColor($name)
    {
        return isset(self::$colorsMap[$name]);
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
    public static function getColor($name)
    {
        $defaultColor = 0x7FFF;

        // Return the default color, 0x7FFF, if undef,
        if ($name === '') {
            return $defaultColor;
        }

        // or the color string converted to an integer,
        if (self::isValidColor($name)) {
            return self::$colorsMap[$name];
        }

        // String is unrecognised or arg is outside range,
        if (preg_match("/\D/", $name) || $name > 63) {
            return $defaultColor;
        }

        // or an integer in the valid range
        return $name;
    }
}
