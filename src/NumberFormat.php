<?php

namespace Xls;

class NumberFormat
{
    /**
     * General format
     */
    const TYPE_GENERAL = 0;

    /**
     * Decimal: 0
     */
    const TYPE_DECIMAL_1 = 1;

    /**
     * Decimal: 0.00
     */
    const TYPE_DECIMAL_2 = 2;

    /**
     * Decimal: #,##0
     */
    const TYPE_DECIMAL_3 = 3;

    /**
     * Decimal: #,##0.00
     */
    const TYPE_DECIMAL_4 = 4;

    /**
     * Currency: "$"#,##0_);("$"#,##0)
     */
    const TYPE_CURRENCY_1 = 5;

    /**
     * Currency: "$"#,##0_);[Red]("$"#,##0)
     */
    const TYPE_CURRENCY_2 = 6;

    /**
     * Currency: "$"#,##0.00_);("$"#,##0.00)
     */
    const TYPE_CURRENCY_3 = 7;

    /**
     * Currency: "$"#,##0.00_);[Red]("$"#,##0.00)
     */
    const TYPE_CURRENCY_4 = 8;

    /**
     * Currency: _("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)
     */
    const TYPE_CURRENCY_5 = 41;

    /**
     * Currency: _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
     */
    const TYPE_CURRENCY_6 = 42;

    /**
     * Currency: _("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)
     */
    const TYPE_CURRENCY_7 = 43;

    /**
     * Currency: _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
     */
    const TYPE_CURRENCY_8 = 44;

    /**
     * Percent: 0%
     */
    const TYPE_PERCENT_1 = 9;

    /**
     * Percent: 0.00%
     */
    const TYPE_PERCENT_2 = 10;

    /**
     * Scientific: 0.00E+00
     */
    const TYPE_SCIENTIFIC_1 = 11;

    /**
     * Scientific: ##0.0E+0
     */
    const TYPE_SCIENTIFIC_2 = 48;

    /**
     * Fraction: # ?/?
     */
    const TYPE_FRACTION_1 = 12;

    /**
     * Fraction: # ??/??
     */
    const TYPE_FRACTION_2 = 13;

    /**
     * Date: M/D/YY
     */
    const TYPE_DATE_1 = 14;

    /**
     * Date: D-MMM-YY
     */
    const TYPE_DATE_2 = 15;

    /**
     * Date: D-MMM
     */
    const TYPE_DATE_3 = 16;

    /**
     * Date: MMM-YY
     */
    const TYPE_DATE_4 = 17;

    /**
     * Time: h:mm AM/PM
     */
    const TYPE_TIME_1 = 18;

    /**
     * Time: h:mm:ss AM/PM
     */
    const TYPE_TIME_2 = 19;

    /**
     * Time: h:mm
     */
    const TYPE_TIME_3 = 20;

    /**
     * Time: h:mm:ss
     */
    const TYPE_TIME_4 = 21;

    /**
     * Time: mm:ss
     */
    const TYPE_TIME_5 = 45;

    /**
     * Time: [h]:mm:ss
     */
    const TYPE_TIME_6 = 46;

    /**
     * Time: mm:ss.0
     */
    const TYPE_TIME_7 = 47;

    /**
     * Datetime: M/D/YY h:mm
     */
    const TYPE_DATETIME = 22;

    /**
     * Account: _(#,##0_);(#,##0)
     */
    const TYPE_ACCOUNT_1 = 37;

    /**
     * Account: _(#,##0_);[Red](#,##0)
     */
    const TYPE_ACCOUNT_2 = 38;

    /**
     * Account: _(#,##0.00_);(#,##0.00)
     */
    const TYPE_ACCOUNT_3 = 39;

    /**
     * Account: _(#,##0.00_);[Red](#,##0.00)
     */
    const TYPE_ACCOUNT_4 = 40;

    /**
     * Text: @
     */
    const TYPE_TEXT = 49;

    /**
     * @param $format
     *
     * @return bool
     */
    public static function isBuiltIn($format)
    {
        return preg_match("/^\d+$/", $format) === 1;
    }
}
