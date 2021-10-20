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
     * @var bool If sentence is subscript or not
     */
    public $subscript;

    /**
     * @var int Number of line breaks
     */
    public $br;

    /**
     * @var int Number of tabs
     */
    public $tab;

    /**
     * @var Style|null
     */
    public $style;

    /**
     * @var Insertion|null
     */
    public $insertion;

    /**
     * @var Deletion|null
     */
    public $deletion;

    /**
     * @var string|null
     */
    public $rsidR;

    /**
     * @var string|null
     */
    public $rsidDel;
    /**
     * @var null
     */
    private $hyperLink;


    function __construct(
        $text,
        $bold = false,
        $italic = false,
        $underline = false,
        $br = 0,
        $tab = 0,
        $highlight = false,
        $superscript = false,
        $subscript = false,
        $style = null,
        $insertion = null,
        $deletion = null,
        $rsidR = null,
        $rsidDel = null,
        $hyperLink = null
    ) {
        $this->text = $text;
        $this->bold = $bold;
        $this->italic = $italic;
        $this->underline = $underline;
        $this->br = $br;
        $this->tab = $tab;
        $this->highlight = $highlight;
        $this->superscript = $superscript;
        $this->subscript = $subscript;
        $this->style = $style;
        $this->insertion = $insertion;
        $this->deletion = $deletion;
        $this->rsidR = $rsidR;
        $this->rsidDel = $rsidDel;
        $this->hyperLink = $hyperLink;
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toDocxXML()
    {
        $value = '';
        if($this->hyperLink){
            $value .= $this->hyperLink->toDocxXML();
        }
        if ($this->insertion) {
            $value .= $this->insertion->toDocxXML();
        }
        if ($this->deletion) {
            $value .= $this->deletion->toDocxXML();
        }
        if ($this->rsidR !== null) {
            $value .= '<w:r w:rsidR="' . $this->rsidR . '" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
        } elseif ($this->rsidDel !== null) {
            $value .= '<w:r w:rsidDel="' . $this->rsidDel . '" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
        } else {
            $value .= '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
        }

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
        for ($i = 0; $i < $this->tab; $i++) {
            $value .= "<w:tab/>";
        }

        $value .= '<w:t xml:space="preserve">' . htmlentities($this->text, ENT_XML1) . "</w:t></w:r>";

        if ($this->deletion !== null) {
            $value .= '</w:del>';
        }
        if ($this->insertion !== null) {
            $value .= '</w:ins>';
        }
        if($this->hyperLink !== null) {
            $value .= '</w:hyperlink>';
        }

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
     * @param bool $firstWrappedInFont
     * @param bool $lastWrappedInBold
     * @param bool $lastWrappedInItalic
     * @param bool $lastWrappedInUnderline
     * @param bool $lastWrappedInHighlight
     * @param bool $lastWrappedInSuperscript
     * @param bool $lastWrappedInSubscript
     * @param bool $lastWrappedInFont
     * @return string HTML string
     */
    public function toHTML(
        $firstWrappedInBold = true,
        $firstWrappedInItalic = true,
        $firstWrappedInUnderline = true,
        $firstWrappedInHighlight = true,
        $firstWrappedInSuperscript = true,
        $firstWrappedInSubscript = true,
        $firstWrappedInFont = true,
        $lastWrappedInBold = true,
        $lastWrappedInItalic = true,
        $lastWrappedInUnderline = true,
        $lastWrappedInHighlight = true,
        $lastWrappedInSuperscript = true,
        $lastWrappedInSubscript = true,
        $lastWrappedInFont = true
    ) {
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
        if ($this->style !== null && !$this->style->isEmpty() && $firstWrappedInFont) {
            $value .= "<font>";
        }

        $value .= htmlentities($this->text);

        if ($this->style !== null && !$this->style->isEmpty() && $lastWrappedInFont) {
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