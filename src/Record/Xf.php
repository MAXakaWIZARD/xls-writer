<?php
namespace Xls\Record;

use Xls\Format as XlsFormat;
use Xls\NumberFormat;

class Xf extends AbstractRecord
{
    const NAME = 'XF';
    const ID = 0x00E0;

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
        $style = $this->getStyle($style, $format);

        $this->checkBorders($format);

        $border1 = $this->getBorder1($format);
        $border2 = $this->getBorder2($format);

        $icv = $format->fgColor; // fg and bg pattern colors
        $icv |= $format->bgColor << 7;

        $options = 0x00;
        $data = pack(
            "vvvC",
            $format->getFont()->getIndex(),
            $format->getNumFormatIndex(),
            $style,
            $this->getAlignment($format)
        );
        $data .= pack("CCC", $format->rotation, $options, $this->getUsedAttr($format));
        $data .= pack("VVv", $border1, $border2, $icv);

        return $this->getFullRecord($data);
    }

    /**
     * @param XlsFormat $format
     *
     * @return int
     */
    protected function getAlignment($format)
    {
        $align = $format->textHorAlign;
        $align |= $format->textWrap << 3;
        $align |= $format->textVertAlign << 4;

        $textJustlast = 0x00;
        $align |= $textJustlast << 7;

        return $align;
    }

    /**
     * @param XlsFormat $format
     *
     * @return int
     */
    protected function getUsedAttr($format)
    {
        $flags = $this->getFlags($format);

        $usedAttr = $flags['Num'] << 2;
        $usedAttr |= $flags['Fnt'] << 3;
        $usedAttr |= $flags['Alc'] << 4;
        $usedAttr |= $flags['Bdr'] << 5;
        $usedAttr |= $flags['Pat'] << 6;
        $usedAttr |= $flags['Prot'] << 7;

        return $usedAttr;
    }

    /**
     * @param XlsFormat $format
     *
     * @return int
     */
    protected function getBorder1($format)
    {
        $border1 = $format->getBorderStyle('left'); // Border line style and color
        $border1 |= $format->getBorderStyle('right') << 4;
        $border1 |= $format->getBorderStyle('top') << 8;
        $border1 |= $format->getBorderStyle('bottom') << 12;
        $border1 |= $format->getBorderColor('left') << 16;
        $border1 |= $format->getBorderColor('right') << 23;
        $diagTlToRb = 0;
        $diagTrToLb = 0;
        $border1 |= $diagTlToRb << 30;
        $border1 |= $diagTrToLb << 31;

        return $border1;
    }

    /**
     * @param XlsFormat $format
     *
     * @return int
     */
    protected function getBorder2($format)
    {
        $border2 = $format->getBorderColor('top'); // Border color
        $border2 |= $format->getBorderColor('bottom') << 7;
        $border2 |= $format->diagColor << 14;
        $border2 |= $format->diag << 21;
        $border2 |= $format->pattern << 26;

        return $border2;
    }

    /**
     * @param string $style
     * @param XlsFormat $format
     *
     * @return int
     */
    protected function getStyle($style, $format)
    {
        // Set the type of the XF record and some of the attributes.
        if ($style == 'style') {
            $style = 0xFFF5;
        } else {
            $style = $format->locked;
            $style |= $format->hidden << 1;
        }

        return $style;
    }

    /**
     * Zero border colors if no borders set
     * @param XlsFormat $format
     */
    protected function checkBorders($format)
    {
        if ($format->diag == 0) {
            $format->diagColor = 0;
        }
    }

    /**
     * @param XlsFormat $format
     *
     * @return array
     */
    protected function getFlags($format)
    {
        return array(
            'Num' => ($format->getNumFormat() != NumberFormat::TYPE_GENERAL) ? 1 : 0,
            'Fnt' => ($format->getFont()->getIndex() != 0) ? 1 : 0,
            'Alc' => ($format->textWrap) ? 1 : 0,
            'Bdr' => ($format->getBorderStyle('top')
                || $format->getBorderStyle('right')
                || $format->getBorderStyle('bottom')
                || $format->getBorderStyle('left')) ? 1 : 0,
            'Pat' => ($format->fgColor != 0x40
                || $format->bgColor != 0x41
                || $format->pattern) ? 1 : 0,
            'Prot' => $format->locked | $format->hidden
        );
    }
}
