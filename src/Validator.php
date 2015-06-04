<?php

namespace Xls;

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

    const TYPE_ANY = 0x00;
    const TYPE_INTEGER = 0x01;
    const TYPE_DECIMAL = 0x02;
    const TYPE_USER_LIST = 0x03;
    const TYPE_DATE = 0x04;
    const TYPE_TIME = 0x05;
    const TYPE_TEXT_LENGTH = 0x06;
    const TYPE_FORMULA = 0x07;

    const ERROR_STOP = 0x00;
    const ERROR_WARNING = 0x01;
    const ERROR_INFO = 0x02;

    protected $dataType = self::TYPE_INTEGER;
    protected $errorStyle = self::ERROR_STOP;
    protected $allowBlank = false;
    protected $showDropDown = false;
    protected $showPrompt = false;
    protected $showError = true;
    protected $titlePrompt = "\x00";
    protected $descrPrompt = "\x00";
    protected $titleError = "\x00";
    protected $descrError = "\x00";
    protected $operator = self::OP_BETWEEN;
    protected $formula1 = '';
    protected $formula2 = '';

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
    }

    /**
     * @param int $operator
     */
    public function setOperator($operator)
    {
        $this->operator = $operator;
    }

    /**
     * @param int $dataType
     */
    public function setDataType($dataType)
    {
        $this->dataType = $dataType;
    }

    /**
     * @param string $promptTitle
     * @param string $promptDescription
     * @param bool   $showPrompt
     */
    public function setPrompt($promptTitle = "\x00", $promptDescription = "\x00", $showPrompt = true)
    {
        $this->showPrompt = $showPrompt;
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
        $this->showError = $showError;
        $this->titleError = $errorTitle;
        $this->descrError = $errorDescription;
    }

    /**
     *
     */
    public function allowBlank()
    {
        $this->allowBlank = true;
    }

    /**
     *
     */
    public function onInvalidStop()
    {
        $this->errorStyle = self::ERROR_STOP;
    }

    /**
     *
     */
    public function onInvalidWarn()
    {
        $this->errorStyle = self::ERROR_WARNING;
    }

    /**
     *
     */
    public function onInvalidInfo()
    {
        $this->errorStyle = self::ERROR_INFO;
    }

    /**
     * @param $formula
     *
     */
    public function setFormula1($formula)
    {
        $this->formula1 = $formula;
    }

    /**
     * @param $formula
     *
     */
    public function setFormula2($formula)
    {
        $this->formula2 = $formula;
    }

    /**
     * @return int
     */
    public function getOptions()
    {
        $options = 0x00;

        $options |= $this->dataType;
        $options |= $this->errorStyle << 4;

        if ($this->dataType === self::TYPE_USER_LIST
            && preg_match('/^\".*\"$/', $this->formula1)
        ) {
            //explicit list options, separated by comma
            $options |= 0x01 << 7;
        }

        $options |= intval($this->allowBlank) << 8;
        $options |= intval(!$this->showDropDown) << 9;
        $options |= intval($this->showPrompt) << 18;
        $options |= intval($this->showError) << 19;
        $options |= $this->operator << 20;

        return $options;
    }

    /**
     * @param boolean $showDropDown
     */
    public function setShowDropDown($showDropDown = true)
    {
        $this->showDropDown = $showDropDown;
    }

    /**
     * @param Range $range
     *
     * @return string
     */
    public function getData(Range $range)
    {
        $data = pack("V", $this->getOptions());

        $data .= pack("vC", strlen($this->titlePrompt), 0x00) . $this->titlePrompt;
        $data .= pack("vC", strlen($this->titleError), 0x00) . $this->titleError;
        $data .= pack("vC", strlen($this->descrPrompt), 0x00) . $this->descrPrompt;
        $data .= pack("vC", strlen($this->descrError), 0x00) . $this->descrError;

        $data .= $this->packFormula($this->formula1);
        $data .= $this->packFormula($this->formula2);

        $data .= \Xls\Subrecord\Range::getData(array($range));

        return $data;
    }

    protected function packFormula($formula)
    {
        if ($this->dataType === self::TYPE_USER_LIST) {
            $formula = str_replace(',', chr(0), $formula);
        }

        if ($formula != '') {
            $formula = $this->formulaParser->getReversePolish($formula);
        }

        return pack("vv", strlen($formula), 0x00) . $formula;
    }
}
