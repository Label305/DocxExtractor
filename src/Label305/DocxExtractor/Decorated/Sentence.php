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

    public $bold;

    public $italic;

    public $underline;

    public $br;

    function __construct($text, $bold, $italic, $underline, $br)
    {
        $this->text = $text;
        $this->bold = $bold;
        $this->italic = $italic;
        $this->underline = $underline;
        $this->br = $br;
    }

    /**
     * To docx xml string
     *
     * @return sting
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
            $value .= "<w:i/>";
        }

        $value .= "</w:rPr>";

        for ($i = 0; $i < $this->br; $i++) {
            $value .= "<w:br/>";
        }

        $value .= "<w:t>" . htmlentities($this->text) . "</w:t></w:r>";

        return $value;
    }


}