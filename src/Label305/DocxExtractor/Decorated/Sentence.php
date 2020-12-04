<?php

namespace Label305\DocxExtractor\Decorated;

/**
 * Class Sentence
 * @package Label305\DocxExtractor\Decorated
 *
 * Represents a <w:r> object in the docx format.
 */
class Sentence {

    /**
     * @var string
     */
    public $text;

    /**
     * @var bool If sentence is bold or not
     */
    public $bold;

    /**
     * @var bool If sentence is italic or not
     */
    public $italic;

    /**
     * @var bool If sentence is underlined or not
     */
    public $underline;

    /**
     * @var bool If sentence is highlighted or not
     */
    public $highlight;

    /**
     * @var bool If sentence is superscript or not
     */
    public $superscript;

    /**
     * @var bool If sentence is subscriot or not
     */
    public $subscript;

    /**
     * @var int Number of line breaks
     */
    public $br;

    /**
     * @var Style|null
     */
    public $style;


    function __construct(
        $text,
        $bold = false,
        $italic = false,
        $underline = false,
        $br = false,
        $highlight = false,
        $superscript = false,
        $subscript = false,
        $style = null
    ) {
        $this->text = $text;
        $this->bold = $bold;
        $this->italic = $italic;
        $this->underline = $underline;
        $this->br = $br;
        $this->highlight = $highlight;
        $this->superscript = $superscript;
        $this->subscript = $subscript;
        $this->style = $style;
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toDocxXML()
    {
        $value = '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
        $value .= '<w:rPr>';
        if ($this->style !== null) {
            $value .= $this->style->toDocxXML();
        }
        if ($this->bold) {
            $value .= "<w:b/>";
        }
        if ($this->italic) {
            $value .= "<w:i/>";
        }
        if ($this->underline) {
            $value .= '<w:u w:val="single"/>';
        }
        if ($this->highlight) {
            if ($this->style !== null && $this->style->highlightColor !== null) {
                $value .= '<w:highlight w:val="' . $this->style->highlightColor . '"/>';
            } else {
                $value .= '<w:highlight w:val="yellow"/>';    
            }
        }
        if ($this->superscript) {
            $value .= '<w:vertAlign w:val="superscript"/>';
        }
        if ($this->subscript) {
            $value .= '<w:vertAlign w:val="subscript"/>';
        }

        $value .= "</w:rPr>";

        for ($i = 0; $i < $this->br; $i++) {
            $value .= "<w:br/>";
        }

        $value .= '<w:t xml:space="preserve">' . htmlentities($this->text, ENT_XML1) . "</w:t></w:r>";

        return $value;
    }

    /**
     * Convert to HTML
     *
     * To prevent duplicate tags (e.g. <strong) and allow wrapping you can use the parameters. If they are set to false
     * a tag will not be opened or closed.
     *
     * @param bool $firstWrappedInBold
     * @param bool $firstWrappedInItalic
     * @param bool $firstWrappedInUnderline
     * @param bool $firstWrappedInHighlight
     * @param bool $firstWrappedInSuperscript
     * @param bool $firstWrappedInSubscript
     * @param bool $lastWrappedInBold
     * @param bool $lastWrappedInItalic
     * @param bool $lastWrappedInUnderline
     * @param bool $lastWrappedInHighlight
     * @param bool $lastWrappedInSuperscript
     * @param bool $lastWrappedInSubscript
     * @return string HTML string
     */
    public function toHTML(
        $firstWrappedInBold = true,
        $firstWrappedInItalic = true,
        $firstWrappedInUnderline = true,
        $firstWrappedInHighlight = true,
        $firstWrappedInSuperscript = true,
        $firstWrappedInSubscript = true,
        $lastWrappedInBold = true,
        $lastWrappedInItalic = true,
        $lastWrappedInUnderline = true,
        $lastWrappedInHighlight = true,
        $lastWrappedInSuperscript = true,
        $lastWrappedInSubscript = true
    )
    {
        $value = '';

        for ($i = 0; $i < $this->br; $i++) {
            $value .= "<br />";
        }

        if ($this->highlight && $firstWrappedInHighlight) {
            $value .= "<mark>";
        }
        if ($this->bold && $firstWrappedInBold) {
            $value .= "<strong>";
        }
        if ($this->italic && $firstWrappedInItalic) {
            $value .= "<em>";
        }
        if ($this->underline && $firstWrappedInUnderline) {
            $value .= "<u>";
        }
        if ($this->subscript && $firstWrappedInSubscript) {
            $value .= "<sub>";
        }
        if ($this->superscript && $firstWrappedInSuperscript) {
            $value .= "<sup>";
        }
        if ($this->style !== null &&
            !$this->highlight &&
            !$this->bold &&
            !$this->italic &&
            !$this->underline &&
            !$this->subscript &&
            !$this->superscript
        ) {
            $value .= "<font>";
        }

        $value .= htmlentities($this->text);

        if ($this->style !== null &&
            !$this->highlight &&
            !$this->bold &&
            !$this->italic &&
            !$this->underline &&
            !$this->subscript &&
            !$this->superscript
        ) {
            $value .= "</font>";
        }
        if ($this->superscript && $lastWrappedInSuperscript) {
            $value .= "</sup>";
        }
        if ($this->subscript && $lastWrappedInSubscript) {
            $value .= "</sub>";
        }
        if ($this->underline && $lastWrappedInUnderline) {
            $value .= "</u>";
        }
        if ($this->italic && $lastWrappedInItalic) {
            $value .= "</em>";
        }
        if ($this->bold && $lastWrappedInBold) {
            $value .= "</strong>";
        }
        if ($this->highlight && $lastWrappedInHighlight) {
            $value .= "</mark>";
        }

        return $value;
    }
}