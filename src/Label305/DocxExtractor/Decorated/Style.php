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
        return empty($this->rFonts) &&
            empty($this->color) &&
            empty($this->lang) &&
            empty($this->sz) &&
            empty($this->szCs) &&
            empty($this->position) &&
            empty($this->spacing) &&
            empty($this->highlightColor);
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toDocxXML()
    {
        $properties = [
            'color',
            'lang',
            'sz',
            'szCs',
            'position',
            'spacing',
        ];

        $value = '';
        if ($this->rFonts !== null) {
            $value .= '<w:rFonts w:ascii="' . $this->rFonts . '" w:hAnsi="' . $this->rFonts . '" w:cs="' . $this->rFonts . '"/>';
        }
        foreach ($properties as $property) {
            if ($this->$property !== null) {
                $value .= '<w:' . $property . ' w:val="' . $this->$property . '"/>';
            }
        }
        return $value;
    }
}