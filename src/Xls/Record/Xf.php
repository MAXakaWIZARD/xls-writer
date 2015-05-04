<?php

namespace Xls\Record;

use Xls\Format as XlsFormat;
use Xls\Biff5;

class Xf extends AbstractRecord
{
    const NAME = 'XF';
    const ID = 0x00E0;
    const LENGTH = 0x00;

    /**
     * Generate an Excel BIFF XF record.
     *
     * @param XlsFormat $format
     * @param string $style The type of the XF record ('style' or 'cell').
     *
     * @return string
     */
    public function getData($format, $style)
    {
        // Set the type of the XF record and some of the attributes.
        if ($style == 'style') {
            $style = 0xFFF5;
        } else {
            $style = $format->locked;
            $style |= $format->hidden << 1;
        }

        // Flags to indicate if attributes have been set.
        $atrNum = ($format->numFormat != 0) ? 1 : 0;
        $atrFnt = ($format->fontIndex != 0) ? 1 : 0;
        $atrAlc = ($format->textWrap) ? 1 : 0;
        $atrBdr = ($format->bottom
            || $format->top
            || $format->left
            || $format->right) ? 1 : 0;
        $atrPat = (($format->fgColor != 0x40)
            || ($format->bgColor != 0x41)
            || $format->pattern) ? 1 : 0;
        $atrProt = $format->locked | $format->hidden;

        // Zero the default border colour if the border has not been set.
        if ($format->bottom == 0) {
            $format->bottomColor = 0;
        }

        if ($format->top == 0) {
            $format->topColor = 0;
        }

        if ($format->right == 0) {
            $format->rightColor = 0;
        }

        if ($format->left == 0) {
            $format->leftColor = 0;
        }

        if ($format->diag == 0) {
            $format->diagColor = 0;
        }

        $ifnt = $format->fontIndex; // Index to FONT record
        $ifmt = $format->numFormat; // Index to FORMAT record
        if ($format->getVersion() === Biff5::VERSION) {
            $length = 0x0010;

            $align = $format->textHorAlign; // Alignment
            $align |= $format->textWrap << 3;
            $align |= $format->textVertAlign << 4;
            $align |= $format->textJustlast << 7;
            $align |= $format->rotation << 8;
            $align |= $atrNum << 10;
            $align |= $atrFnt << 11;
            $align |= $atrAlc << 12;
            $align |= $atrBdr << 13;
            $align |= $atrPat << 14;
            $align |= $atrProt << 15;

            $icv = $format->fgColor; // fg and bg pattern colors
            $icv |= $format->bgColor << 7;

            $fill = $format->pattern; // Fill and border line style
            $fill |= $format->bottom << 6;
            $fill |= $format->bottomColor << 9;

            $border1 = $format->top; // Border line style and color
            $border1 |= $format->left << 3;
            $border1 |= $format->right << 6;
            $border1 |= $format->topColor << 9;

            $border2 = $format->leftColor; // Border color
            $border2 |= $format->rightColor << 7;

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
            $length = 0x0014;

            $align = $format->textHorAlign; // Alignment
            $align |= $format->textWrap << 3;
            $align |= $format->textVertAlign << 4;
            $align |= $format->textJustlast << 7;

            $usedAttr = $atrNum << 2;
            $usedAttr |= $atrFnt << 3;
            $usedAttr |= $atrAlc << 4;
            $usedAttr |= $atrBdr << 5;
            $usedAttr |= $atrPat << 6;
            $usedAttr |= $atrProt << 7;

            $icv = $format->fgColor; // fg and bg pattern colors
            $icv |= $format->bgColor << 7;

            $border1 = $format->left; // Border line style and color
            $border1 |= $format->right << 4;
            $border1 |= $format->top << 8;
            $border1 |= $format->bottom << 12;
            $border1 |= $format->leftColor << 16;
            $border1 |= $format->rightColor << 23;
            $diagTlToRb = 0;
            $diagTrToLb = 0;
            $border1 |= $diagTlToRb << 30;
            $border1 |= $diagTrToLb << 31;

            $border2 = $format->topColor; // Border color
            $border2 |= $format->bottomColor << 7;
            $border2 |= $format->diagColor << 14;
            $border2 |= $format->diag << 21;
            $border2 |= $format->pattern << 26;

            $rotation = $format->rotation;
            $biff8Options = 0x00;
            $data = pack("vvvC", $ifnt, $ifmt, $style, $align);
            $data .= pack("CCC", $rotation, $biff8Options, $usedAttr);
            $data .= pack("VVv", $border1, $border2, $icv);
        }

        return $this->getHeader($length) . $data;
    }
}
