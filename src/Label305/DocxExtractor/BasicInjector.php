<?php namespace Label305\DocxExtractor;

use DOMDocument;
use DOMNode;
use DOMText;

class BasicInjector extends DocxHandler implements Injector {

    /**
     * @param $mapping
     * @param $fileToInjectLocationHandle
     * @param $saveLocationHandle
     * @throws DocxFileException
     * @throws DocxParsingException
     * @return void
     */
    public function injectMappingAndCreateNewFile($mapping, $fileToInjectLocationHandle, $saveLocationHandle)
    {
        $prepared = $this->prepareDocumentForReading($fileToInjectLocationHandle);

        $this->assignMappedValues($prepared['dom']->documentElement, $mapping);

        $this->saveDocument($prepared['dom'], $prepared["archive"], $saveLocationHandle);
    }

    /**
     * @param DOMNode $node
     * @param $mapping
     */
    protected function assignMappedValues(DOMNode $node, $mapping)
    {
        if ($node instanceof DOMText) {
            $results = [];
            preg_match("/%[0-9]*%/", $node->nodeValue, $results);

            if (count($results) > 0) {
                $key = trim($results[0], '%');
                $node->nodeValue = $mapping[$key];
            }
        }

        if ($node->childNodes !== null) {
            foreach ($node->childNodes as $child) {
                $this->assignMappedValues($child, $mapping);
            }
        }
    }
}