<?php

namespace Label305\DocxExtractor\Decorated;

/**
 * Class Hyperlink
 * @package Label305\DocxExtractor\Decorated
 *
 * Represents the style contents of a <w:hyperlink> object in the docx format.
 */
class Hyperlink
{

    /**
     * @var string|null
     */
    public $id;
    /**
     * @var string|null
     */
    public $tgtFrame;
    /**
     * @var string|null
     */
    public $history = "1";

    /**
     * @param string|null $id
     * @param string|null $tgtFrame
     * @param string|null $history
     */
    public function __construct(?string $id, ?string $tgtFrame, ?string $history)
    {
        $this->id = $id;
        $this->tgtFrame = $tgtFrame;
        $this->history = $history;
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toDocxXML(): string
    {
        $properties = [
            'tgtFrame',
            'history'
        ];
        $value = '<w:hyperlink r:id="' . $this->id . '" ';
        foreach ($properties as $property) {
            if ($this->$property !== null) {
                $value .= ' w:' . $property . '="' . $this->$property . '" ';
            }
        }
        $value .= ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
        $value .= ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';

        return $value;
    }
}