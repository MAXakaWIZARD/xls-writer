<?php

function pearIsError($data, $code = null)
{
    if (!is_a($data, 'PEAR_Error')) {
        return false;
    }
    if (is_null($code)) {
        return true;
    } elseif (is_string($code)) {
        return $data->getMessage() == $code;
    }
    return $data->getCode() == $code;
}
