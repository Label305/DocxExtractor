<?php

namespace Label305\DocxExtractor\Decorated;

/**
 * Class Deletion
 * @package Label305\DocxExtractor\Decorated
 *
 * Represents the style contents of a <w:ins> object in the docx format.
 */
class Deletion {

    /**
     * @var string|null
     */
    public $id;
    /**
     * @var string|null
     */
    public $author;
    /**
     * @var string|null
     */
    public $date;

    function __construct($id, $author, $date) {
        $this->id = $id;
        $this->author = $author;
        $this->date = $date;
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toDocxXML()
    {
        $value = '<w:del';

        if ($this->id !== null) {
            $value .= ' w:id="' . $this->id . '" ';
        }
        if ($this->author !== null) {
            $value .= ' w:author="' . $this->author . '" ';
        }
        if ($this->date !== null) {
            $value .= ' w:date="' . $this->date . '" ';
        }
        $value .= ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';

        return $value;
    }
}