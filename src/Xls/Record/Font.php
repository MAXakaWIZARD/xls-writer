<?php

namespace Xls\Record;

use Xls\Format as XlsFormat;
use Xls\Biff5;

class Font extends AbstractRecord
{
    const NAME = 'FONT';
    const ID = 0x0031;
    const LENGTH = 0x00;

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
        $encoding = 0;

        $cch = strlen($format->fontName); // Length of font name
        if ($format->getVersion() === Biff5::VERSION) {
            $length = 0x0F + $cch; // Record length
        } else {
            $length = 0x10 + $cch;
        }

        $reserved = 0x00; // Reserved

        $grbit = 0x00; // Font attributes
        if ($format->italic) {
            $grbit |= 0x02;
        }
        if ($format->fontStrikeout) {
            $grbit |= 0x08;
        }
        if ($format->fontOutline) {
            $grbit |= 0x10;
        }
        if ($format->fontShadow) {
            $grbit |= 0x20;
        }

        if ($format->getVersion() === Biff5::VERSION) {
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

        return $this->getHeader($length) . $data . $format->fontName;
    }
}
