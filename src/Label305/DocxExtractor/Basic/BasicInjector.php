<?php namespace Label305\DocxExtractor\Basic;

use DOMDocument;
use DOMNode;
use DOMText;
use Label305\DocxExtractor\Decorated\Paragraph;
use Label305\DocxExtractor\DocxFileException;
use Label305\DocxExtractor\DocxHandler;
use Label305\DocxExtractor\DocxParsingException;
use Label305\DocxExtractor\Injector;

class BasicInjector extends DocxHandler implements Injector {

    /**
     * @param Paragraph[]|array $mapping
     * @param string $fileToInjectLocationPath
     * @param string $saveLocationPath
     * @throws DocxFileException
     * @throws DocxParsingException
     * @return void
     */
    public function injectMappingAndCreateNewFile(array $mapping, string $fileToInjectLocationPath, string $saveLocationPath): void
    {
        $prepared = $this->prepareDocumentForReading($fileToInjectLocationPath);

        $this->assignMappedValues($prepared['dom']->documentElement, $mapping);

        $this->saveDocument($prepared['dom'], $prepared["archive"], $saveLocationPath);
    }

    /**
     * @param DOMNode $node
     * @param $mapping
     */
    protected function assignMappedValues(DOMNode $node, $mapping): void
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