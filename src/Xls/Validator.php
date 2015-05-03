<?php

namespace Xls;

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
    protected $titlePrompt;
    protected $descrPrompt;
    protected $titleError;
    protected $descrError;
    protected $operator;
    protected $formula1;
    protected $formula2;

    /**
     * The parser from the workbook. Used to parse validation formulas also
     *
     * @var FormulaParser
     */
    protected $formulaParser;

    /**
     * @param FormulaParser $formulaParser
     */
    public function __construct(FormulaParser $formulaParser)
    {
        $this->formulaParser = $formulaParser;
        $this->type = 0x01;
        $this->style = 0x00;
        $this->fixedList = false;
        $this->blank = false;
        $this->incell = false;
        $this->showprompt = false;
        $this->showerror = true;
        $this->titlePrompt = "\x00";
        $this->descrPrompt = "\x00";
        $this->titleError = "\x00";
        $this->descrError = "\x00";
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
        $this->titlePrompt = $promptTitle;
        $this->descrPrompt = $promptDescription;
    }

    /**
     * @param string $errorTitle
     * @param string $errorDescription
     * @param bool   $showError
     */
    public function setError($errorTitle = "\x00", $errorDescription = "\x00", $showError = true)
    {
        $this->showerror = $showError;
        $this->titleError = $errorTitle;
        $this->descrError = $errorDescription;
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
        $this->formulaParser->parse($formula);

        $this->formula1 = $this->formulaParser->toReversePolish();

        return true;
    }

    /**
     * @param $formula
     *
     * @return bool|string
     */
    public function setFormula2($formula)
    {
        $this->formulaParser->parse($formula);

        $this->formula2 = $this->formulaParser->toReversePolish();

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
        $titlePromptLen = strlen($this->titlePrompt);
        $descrPromptLen = strlen($this->descrPrompt);
        $titleErrorLen = strlen($this->titleError);
        $descrErrorLen = strlen($this->descrError);

        $formula1Size = strlen($this->formula1);
        $formula2Size = strlen($this->formula2);

        $data = pack("V", $this->getOptions());
        $data .= pack("vC", $titlePromptLen, 0x00) . $this->titlePrompt;
        $data .= pack("vC", $titleErrorLen, 0x00) . $this->titleError;
        $data .= pack("vC", $descrPromptLen, 0x00) . $this->descrPrompt;
        $data .= pack("vC", $descrErrorLen, 0x00) . $this->descrError;

        $data .= pack("vv", $formula1Size, 0x0000) . $this->formula1;
        $data .= pack("vv", $formula2Size, 0x0000) . $this->formula2;

        return $data;
    }
}
