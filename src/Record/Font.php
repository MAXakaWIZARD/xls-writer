<?php
namespace Xls\Record;

use Xls\Font as XlsFont;

class Font extends AbstractRecord
{
    const NAME = 'FONT';
    const ID = 0x0031;

    /**
     * Generate an Excel BIFF FONT record.
     * @param XlsFont $font
     *
     * @return string
     */
    public function getData(XlsFont $font)
    {
        $dyHeight = $font->getSize() * 20; // Height of font (1/20 of a point)
        $icv = $font->getColor(); // Index to color palette
        $bls = $font->getBold(); // Bold style
        $sss = $font->getScript(); // Superscript/subscript
        $uls = $font->getUnderline(); // Underline
        $bFamily = 0x00; // Font family
        $bCharSet = 0x00; // Character set

        $cch = strlen($font->getName()); // Length of font name

        $reserved = 0x00; // Reserved

        $data = pack(
            "vvvvvCCCCC",
            $dyHeight,
            $this->calcGrbit($font),
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
        $data .= $font->getName();

        return $this->getFullRecord($data);
    }

    /**
     * @param XlsFont $font
     *
     * @return int
     */
    protected function calcGrbit(XlsFont $font)
    {
        $grbit = 0x00;

        $grbit |= intval($font->getItalic()) << 1;
        $grbit |= intval($font->getStrikeout()) << 3;
        $grbit |= intval($font->getOutline()) << 4;
        $grbit |= intval($font->getShadow()) << 5;

        return $grbit;
    }
}
