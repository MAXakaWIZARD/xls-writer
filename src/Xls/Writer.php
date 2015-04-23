<?php

namespace Xls;

/**
 * Class for writing Excel Spreadsheets
 *
 */
class Writer extends Workbook
{
    /**
     * Send HTTP headers for the Excel file.
     *
     * @param string $filename The filename to use for HTTP headers
     */
    public function send($filename)
    {
        header("Content-type: application/vnd.ms-excel");
        header("Content-Disposition: attachment; filename=\"$filename\"");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0,pre-check=0");
        header("Pragma: public");
    }
}
