<?php

namespace Label305\DocxExtractor\Decorated;

/**
 * Class Sentence
 * @package Label305\DocxExtractor\Decorated
 *
 * Represents the style contents of a <w:rPr> object in the docx format.
 */
class Style {

    /**
     * @var string|null
     */
    public $rFonts;
    /**
     * @var string|null
     */
    public $color;
    /**
     * @var string|null
     */
    public $lang;
    /**
     * @var string|null
     */
    public $sz;
    /**
     * @var string|null
     */
    public $szCs;
    /**
     * @var string|null
     */
    public $position;
    /**
     * @var string|null
     */
    public $spacing;
    /**
     * @var string|null
     */
    public $highlightColor;

    function __construct($rFonts, $color, $lang, $sz, $szCs, $position, $spacing, $highlightColor) {
        $this->rFonts = $rFonts;
        $this->color = $color;
        $this->lang = $lang;
        $this->sz = $sz;
        $this->szCs = $szCs;
        $this->position = $position;
        $this->spacing = $spacing;
        $this->highlightColor = $highlightColor;
    }

    /**
     * @return bool
     */
    public function isEmpty()
    {
        return ($this->rFonts === null || $this->rFonts === "" ) &&
            ($this->color === null || $this->color === "" ) &&
            ($this->lang === null || $this->lang === "" ) &&
            ($this->sz === null || $this->sz === "" ) &&
            ($this->szCs === null || $this->szCs === "" ) &&
            ($this->position === null || $this->position === "" ) &&
            ($this->spacing === null || $this->spacing === "" ) &&
            ($this->highlightColor === null || $this->highlightColor === "");
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toDocxXML()
    {
        $value = '';
        if ($this->rFonts !== null) {
            $value .= '<w:rFonts w:ascii="' . $this->rFonts . '" w:hAnsi="' . $this->rFonts . '" w:cs="' . $this->rFonts . '"/>';
        }
        if ($this->color !== null) {
            $value .= '<w:color w:val="' . $this->color . '"/>';
        }
        if ($this->lang !== null) {
            $value .= '<w:lang w:val="' . $this->lang . '"/>';
        }
        if ($this->sz !== null) {
            $value .= '<w:sz w:val="' . $this->sz . '"/>';
        }
        if ($this->szCs !== null) {
            $value .= '<w:szCs w:val="' . $this->szCs . '"/>';
        }
        if ($this->position !== null) {
            $value .= '<w:position w:val="' . $this->position . '"/>';
        }
        if ($this->spacing !== null) {
            $value .= '<w:spacing w:val="' . $this->spacing . '"/>';
        }
        return $value;
    }
}