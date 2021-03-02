<?php

namespace Label305\DocxExtractor\Decorated;

/**
 * Class Insertion
 * @package Label305\DocxExtractor\Decorated
 *
 * Represents the style contents of a <w:ins> object in the docx format.
 */
class Insertion {

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
        $properties = [
            'id',
            'author',
            'date'
        ];
        $value = '<w:ins';
        foreach ($properties as $property) {
            if ($this->$property !== null) {
                $value .= ' w:' . $property . '="' . $this->$property . '" ';
            }
        }
        $value .= ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';

        return $value;
    }
}