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
    private $rFonts;
    /**
     * @var string|null
     */
    private $color;
    /**
     * @var string|null
     */
    private $lang;
    /**
     * @var string|null
     */
    private $sz;
    /**
     * @var string|null
     */
    private $szCs;
    /**
     * @var string|null
     */
    private $position;
    /**
     * @var string|null
     */
    private $spacing;
    /**
     * @var string|null
     */
    private $highlightColor;

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

    /**
     * @return string|null
     */
    public function getRFonts()
    {
        return $this->rFonts;
    }

    /**
     * @return string|null
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * @return string|null
     */
    public function getLang()
    {
        return $this->lang;
    }

    /**
     * @return string|null
     */
    public function getSz()
    {
        return $this->sz;
    }

    /**
     * @return string|null
     */
    public function getSzCs()
    {
        return $this->szCs;
    }

    /**
     * @return string|null
     */
    public function getPosition()
    {
        return $this->position;
    }

    /**
     * @return string|null
     */
    public function getSpacing()
    {
        return $this->spacing;
    }

    /**
     * @return string|null
     */
    public function getHighlightColor()
    {
        return $this->highlightColor;
    }
}