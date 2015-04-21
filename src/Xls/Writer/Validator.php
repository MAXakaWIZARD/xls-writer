<?php
/*
*  Module written by Herman Kuiper <herman@ozuzo.net>
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
 * Baseclass for generating Excel DV records (validations)
 *
 * @author   Herman Kuiper
 * @category FileFormats
 * @package  Spreadsheet_Excel_Writer
 */
class Validator
{
    const OP_BETWEEN = 0x00;
    const OP_NOTBETWEEN = 0x01;
    const OP_EQUAL = 0x02;
    const OP_NOTEQUAL = 0x03;
    const OP_GT = 0x04;
    const OP_LT = 0x05;
    const OP_GTE = 0x06;
    const OP_LTE = 0x07;

    protected $type;
    protected $style;
    protected $fixedList;
    protected $blank;
    protected $incell;
    protected $showprompt;
    protected $showerror;
    protected $title_prompt;
    protected $descr_prompt;
    protected $title_error;
    protected $descr_error;
    protected $operator;
    protected $formula1;
    protected $formula2;

    /**
     * The parser from the workbook. Used to parse validation formulas also
     * @var Parser
     */
    protected $parser;

    /**
     * @param $parser
     */
    public function __construct($parser)
    {
        $this->parser = $parser;
        $this->type = 0x01; // FIXME: add method for setting datatype
        $this->style = 0x00;
        $this->fixedList = false;
        $this->blank = false;
        $this->incell = false;
        $this->showprompt = false;
        $this->showerror = true;
        $this->title_prompt = "\x00";
        $this->descr_prompt = "\x00";
        $this->title_error = "\x00";
        $this->descr_error = "\x00";
        $this->operator = self::OP_BETWEEN;
        $this->formula1 = '';
        $this->formula2 = '';
    }

    /**
     * @param string $promptTitle
     * @param string $promptDescription
     * @param bool   $showPrompt
     */
    public function setPrompt($promptTitle = "\x00", $promptDescription = "\x00", $showPrompt = true)
    {
        $this->showprompt = $showPrompt;
        $this->title_prompt = $promptTitle;
        $this->descr_prompt = $promptDescription;
    }

    /**
     * @param string $errorTitle
     * @param string $errorDescription
     * @param bool   $showError
     */
    public function setError($errorTitle = "\x00", $errorDescription = "\x00", $showError = true)
    {
        $this->showerror = $showError;
        $this->title_error = $errorTitle;
        $this->descr_error = $errorDescription;
    }

    /**
     *
     */
    public function allowBlank()
    {
        $this->blank = true;
    }

    /**
     *
     */
    public function onInvalidStop()
    {
        $this->style = 0x00;
    }

    /**
     *
     */
    public function onInvalidWarn()
    {
        $this->style = 0x01;
    }

    /**
     *
     */
    public function onInvalidInfo()
    {
        $this->style = 0x02;
    }

    /**
     * @param $formula
     *
     * @return bool|string
     */
    public function setFormula1($formula)
    {
        // Parse the formula using the parser in Parser.php
        $this->parser->parse($formula);

        $this->formula1 = $this->parser->toReversePolish();

        return true;
    }

    /**
     * @param $formula
     *
     * @return bool|string
     */
    public function setFormula2($formula)
    {
        // Parse the formula using the parser in Parser.php
        $this->parser->parse($formula);

        $this->formula2 = $this->parser->toReversePolish();

        return true;
    }

    /**
     * @return int
     */
    public function getOptions()
    {
        $options = $this->type;
        $options |= $this->style << 3;
        if ($this->fixedList) {
            $options |= 0x80;
        }
        if ($this->blank) {
            $options |= 0x100;
        }
        if (!$this->incell) {
            $options |= 0x200;
        }
        if ($this->showprompt) {
            $options |= 0x40000;
        }
        if ($this->showerror) {
            $options |= 0x80000;
        }
        $options |= $this->operator << 20;

        return $options;
    }

    /**
     * @return string
     */
    public function getData()
    {
        $title_prompt_len = strlen($this->title_prompt);
        $descr_prompt_len = strlen($this->descr_prompt);
        $title_error_len = strlen($this->title_error);
        $descr_error_len = strlen($this->descr_error);

        $formula1_size = strlen($this->formula1);
        $formula2_size = strlen($this->formula2);

        $data = pack("V", $this->getOptions());
        $data .= pack("vC", $title_prompt_len, 0x00) . $this->title_prompt;
        $data .= pack("vC", $title_error_len, 0x00) . $this->title_error;
        $data .= pack("vC", $descr_prompt_len, 0x00) . $this->descr_prompt;
        $data .= pack("vC", $descr_error_len, 0x00) . $this->descr_error;

        $data .= pack("vv", $formula1_size, 0x0000) . $this->formula1;
        $data .= pack("vv", $formula2_size, 0x0000) . $this->formula2;

        return $data;
    }
}
