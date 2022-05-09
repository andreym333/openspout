<?php

namespace OpenSpout\Common\Helper\Escaper;

/**
 * Provides functions to escape and unescape data for XLSX files.
 */
class XLSX implements EscaperInterface
{
    /** @var string Regex pattern to detect control characters that need to be escaped */
    private static $escapableControlCharactersPattern;

    /** @var string[] Map containing control characters to be escaped (key) and their escaped value (value) */
    private static $controlCharactersEscapingMap;

    /** @var string[] Map containing control characters to be escaped (value) and their escaped value (key) */
    private static $controlCharactersEscapingReverseMap;

    private static function init()
    {
        // control characters values are from 0 to 1F (hex values) in the ASCII table
        // some characters should not be escaped though: "\t", "\r" and "\n".

        static::$escapableControlCharactersPattern =
            '/([\x00-\x08'.
            // skipping "\t" (0x9) and "\n" (0xA)
            '\x0B-\x0C'.
            // skipping "\r" (0xD)
            '\x0E-\x1F])/';

        static::$controlCharactersEscapingMap = [];
        for ($charValue = 0x00; $charValue <= 0x1F; ++$charValue) {
            $character = \chr($charValue);
            if (preg_match(static::$escapableControlCharactersPattern, $character)) {
                $charHexValue = dechex($charValue);
                $escapedChar = '_x'.sprintf('%04s', strtoupper($charHexValue)).'_';
                static::$controlCharactersEscapingMap[$escapedChar] = $character;
            }
        }

        static::$controlCharactersEscapingReverseMap = array_flip(static::$controlCharactersEscapingMap);
    }

    /**
     * Escapes the given string to make it compatible with XLSX.
     *
     * Excel escapes control characters with _xHHHH_ and also escapes any
     * literal strings of that type by encoding the leading underscore.
     * So "\0" -> _x0000_ and "_x0000_" -> _x005F_x0000_.
     *
     * NOTE: the logic has been adapted from the XlsxWriter library (BSD License)
     *
     * @see https://github.com/jmcnamara/XlsxWriter/blob/f1e610f29/xlsxwriter/sharedstrings.py#L89
     *
     * @param string $string The string to escape
     *
     * @return string The escaped string
     */
    public static function escape($string)
    {
        if (!isset(static::$escapableControlCharactersPattern)) {
            static::init();
        }

        // escapes the escape character: "_x0000_" -> "_x005F_x0000_"
        $escapedString = preg_replace('/_(x[\dA-F]{4})_/', '_x005F_$1_', $string);

        if (preg_match(static::$escapableControlCharactersPattern, $escapedString)) {
            $escapedString = preg_replace_callback(static::$escapableControlCharactersPattern, function ($matches) {
                return static::$controlCharactersEscapingReverseMap[$matches[0]];
            }, $escapedString);
        }

        // @NOTE: Using ENT_QUOTES as XML entities ('<', '>', '&') as well as
        //        single/double quotes (for XML attributes) need to be encoded.
        return htmlspecialchars($escapedString, ENT_QUOTES, 'UTF-8');
    }

    /**
     * Unescapes the given string from XLSX.
     *
     * Excel escapes control characters with _xHHHH_ and also escapes any
     * literal strings of that type by encoding the leading underscore.
     * So "_x0000_" -> "\0" and "_x005F_x0000_" -> "_x0000_"
     *
     * NOTE: the logic has been adapted from the XlsxWriter library (BSD License)
     *
     * @see https://github.com/jmcnamara/XlsxWriter/blob/f1e610f29/xlsxwriter/sharedstrings.py#L89
     *
     * @param string $string The string to unescape
     *
     * @return string The unescaped string
     */
    public static function unescape($string)
    {
        if (!isset(static::$escapableControlCharactersPattern)) {
            static::init();
        }

        // ==============
        // =   WARNING  =
        // ==============
        // It is assumed that the given string has already had its XML entities decoded.
        // This is true if the string is coming from a DOMNode (as DOMNode already decode XML entities on creation).
        // Therefore there is no need to call "htmlspecialchars_decode()".
        $unescapedString = $string;

        foreach (static::$controlCharactersEscapingMap as $escapedCharValue => $charValue) {
            // only unescape characters that don't contain the escaped escape character for now
            $unescapedString = preg_replace("/(?<!_x005F)({$escapedCharValue})/", $charValue, $unescapedString);
        }

        // unescapes the escape character: "_x005F_x0000_" => "_x0000_"
        return preg_replace('/_x005F(_x[\dA-F]{4}_)/', '$1', $unescapedString);
    }
}
