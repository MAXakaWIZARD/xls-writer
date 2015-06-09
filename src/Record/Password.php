<?php

namespace Xls\Record;

class Password extends AbstractRecord
{
    const NAME = 'PASSWORD';
    const ID = 0x0013;

    /**
     * Generate the PASSWORD biff record
     *
     * @param $plaintextPassword
     *
     * @return string
     */
    public function getData($plaintextPassword)
    {
        $data = pack("v", $this->encode($plaintextPassword));

        return $this->getFullRecord($data);
    }

    /**
     * Based on the algorithm provided by Daniel Rentz of OpenOffice.
     * @param string $plaintext The password to be encoded in plaintext.
     * @return string The encoded password
     */
    protected function encode($plaintext)
    {
        $password = 0x0000;
        $pos = 1; // char position

        // split the plain text password in its component characters
        $chars = preg_split('//', $plaintext, -1, PREG_SPLIT_NO_EMPTY);
        foreach ($chars as $char) {
            $value = ord($char) << $pos; // shifted ASCII value
            $rotatedBits = $value >> 15; // rotated bits beyond bit 15
            $value &= 0x7fff; // first 15 bits
            $password ^= ($value | $rotatedBits);
            $pos++;
        }

        $password ^= strlen($plaintext);
        $password ^= 0xCE4B;

        return $password;
    }
}
