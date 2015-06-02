<?php
namespace Xls\Record;

use Xls\Format as XlsFormat;

class Font extends AbstractRecord
{
    const NAME = 'FONT';
    const ID = 0x0031;

    /**
     * Generate an Excel BIFF FONT record.
     * @param XlsFormat $format
     *
     * @return string
     */
    public function getData($format)
    {
        $dyHeight = $format->size * 20; // Height of font (1/20 of a point)
        $icv = $format->color; // Index to color palette
        $bls = $format->bold; // Bold style
        $sss = $format->fontScript; // Superscript/subscript
        $uls = $format->underline; // Underline
        $bFamily = $format->fontFamily; // Font family
        $bCharSet = $format->fontCharset; // Character set

        $cch = strlen($format->fontName); // Length of font name

        $reserved = 0x00; // Reserved

        $data = pack(
            "vvvvvCCCCC",
            $dyHeight,
            $this->calcGrbit($format),
            $icv,
            $bls,
            $sss,
            $uls,
            $bFamily,
            $bCharSet,
            $reserved,
            $cch
        );

        $encoding = 0;
        $data .= pack("C", $encoding);
        $data .= $format->fontName;

        return $this->getFullRecord($data);
    }

    /**
     * @param XlsFormat $format
     *
     * @return int
     */
    protected function calcGrbit(XlsFormat $format)
    {
        $grbit = 0x00;

        $grbit |= intval($format->italic) << 1;
        $grbit |= intval($format->fontStrikeout) << 3;
        $grbit |= intval($format->fontOutline) << 4;
        $grbit |= intval($format->fontShadow) << 5;

        return $grbit;
    }
}
