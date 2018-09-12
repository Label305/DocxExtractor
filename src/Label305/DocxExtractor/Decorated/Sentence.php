<?php namespace Label305\DocxExtractor\Decorated;
use DOMDocument;
use DOMDocumentFragment;

/**
 * Class Sentence
 * @package Label305\DocxExtractor\Decorated
 *
 * Represents a <w:r> object in the docx format.
 */
class Sentence {

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
     * @var int Number of line breaks
     */
    public $br;

    function __construct($text, $bold, $italic, $underline, $br, $highlight)
    {
        $this->text = $text;
        $this->bold = $bold;
        $this->italic = $italic;
        $this->underline = $underline;
        $this->br = $br;
        $this->highlight = $highlight;
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toDocxXML()
    {
        $value = '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr>';

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
            $value .= '<w:highlight w:val="yellow"/>';
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
     * @param bool $lastWrappedInBold
     * @param bool $lastWrappedInItalic
     * @param bool $lastWrappedInUnderline
     * @param bool $firstWrappedInHighlight
     * @param bool $lastWrappedInHighlight
     * @return string HTML string
     */
    public function toHTML(
        $firstWrappedInBold = true,
        $firstWrappedInItalic = true,
        $firstWrappedInUnderline = true,
        $firstWrappedInHighlight = true,
        $lastWrappedInBold = true,
        $lastWrappedInItalic = true,
        $lastWrappedInUnderline = true,
        $lastWrappedInHighlight = true
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

        $value .= htmlentities($this->text);

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