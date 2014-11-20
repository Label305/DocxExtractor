<?php
/**
 * Created by PhpStorm.
 * User: Thijs
 * Date: 13-11-14
 * Time: 11:46
 */

namespace Label305\DocxExtractor\Decorated;


use DOMNode;
use DOMText;
use Label305\DocxExtractor\DocxFileException;
use Label305\DocxExtractor\DocxHandler;
use Label305\DocxExtractor\DocxParsingException;
use Label305\DocxExtractor\Injector;

class DecoratedTextInjector extends DocxHandler implements Injector {

    /**
     * @param $mapping
     * @param $fileToInjectLocationPath
     * @param $saveLocationPath
     * @throws DocxFileException
     * @throws DocxParsingException
     * @return void
     */
    public function injectMappingAndCreateNewFile($mapping, $fileToInjectLocationPath, $saveLocationPath)
    {
        $prepared = $this->prepareDocumentForReading($fileToInjectLocationPath);

        $this->assignMappedValues($prepared['dom']->documentElement, $mapping);

        $this->saveDocument($prepared['dom'], $prepared["archive"], $saveLocationPath);
    }

    /**
     * @param DOMNode $node
     * @param array $mapping should be a list of Paragraph objects
     */
    protected function assignMappedValues(DOMNode $node, $mapping)
    {
        if ($node instanceof DOMText) {
            $results = [];
            preg_match("/%[0-9]*%/", $node->nodeValue, $results);

            if (count($results) > 0) {
                $key = trim($results[0], '%');

                $parent = $node->parentNode;

                foreach ($mapping[$key] as $sentence) {

                    $fragment = $parent->ownerDocument->createDocumentFragment();

                    $fragment->appendXML($sentence->toDocxXML());

                    $parent->insertBefore($fragment, $node);
                }

                $parent->removeChild($node);
            }
        }

        if ($node->childNodes !== null) {
            foreach ($node->childNodes as $child) {
                $this->assignMappedValues($child, $mapping);
            }
        }
    }
}