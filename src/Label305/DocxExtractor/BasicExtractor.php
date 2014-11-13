<?php namespace Label305\DocxExtractor;


use DOMDocument;
use DOMNode;
use DOMText;


class BasicExtractor extends DocxHandler implements Extractor {

    /**
     * @var int
     */
    protected $nextTagIdentifier;

    /**
     * @param $originalFileHandle
     * @param $mappingFileSaveLocationHandle
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return Array The mapping of all the strings
     */
    public function extractStringsAndCreateMappingFile($originalFileHandle, $mappingFileSaveLocationHandle)
    {
        $prepared = $this->prepareDocumentForReading($originalFileHandle);

        $this->nextTagIdentifier = 0;
        $result = $this->replaceAndMapValues($prepared['dom']->documentElement);

        $this->saveDocument($prepared['dom'], $prepared["archive"], $mappingFileSaveLocationHandle);

        return $result;
    }

    /**
     * Override this method to make a more complex replace and mapping
     *
     * @param DOMNode $node
     * @return array returns the mapping array
     */
    protected function replaceAndMapValues(DOMNode $node)
    {
        $result = [];

        if ($node instanceof DOMText) {
            $result[$this->nextTagIdentifier] = $node->nodeValue;
            $node->nodeValue = "%".$this->nextTagIdentifier."%";
            $this->nextTagIdentifier++;
        }

        if ($node->childNodes !== null) {
            foreach ($node->childNodes as $child) {
                $result = array_merge(
                    $result,
                    $this->replaceAndMapValues($child)
                );
            }
        }

        return $result;
    }

}