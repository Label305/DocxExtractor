<?php

namespace Label305\DocxExtractor\Decorated;

use DOMNode;
use DOMText;
use Label305\DocxExtractor\DocxFileException;
use Label305\DocxExtractor\DocxHandler;
use Label305\DocxExtractor\DocxParsingException;
use Label305\DocxExtractor\Injector;

class DecoratedTextInjector extends DocxHandler implements Injector {

    /**
     * @var string|null
     */
    private $direction;

    /**
     * @param string|null $direction
     * @throws \Exception
     */
    public function setDirection($direction) {
        if (!in_array($direction, ['ltr', 'rtl'])) {
            throw new \Exception('Direction should be ltr or rtl');
        }
        $this->direction = $direction;
    }

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

                if (isset($mapping[$key])) {
                    $parent = $node->parentNode;

                    if ($this->direction !== null) {
                        $styleNode = $this->addOrFindParagraphStyleNode($parent);
                        $this->addParagraphDirection($styleNode);
                    }
                    foreach ($mapping[$key] as $sentence) {
                        $fragment = $parent->ownerDocument->createDocumentFragment();
                        $fragment->appendXML($sentence->toDocxXML());
                        $parent->insertBefore($fragment, $node);
                    }
                    $parent->removeChild($node);
                }
            }
        }

        if ($node->childNodes !== null) {
            foreach ($node->childNodes as $child) {
                $this->assignMappedValues($child, $mapping);
            }
        }
    }

    private function addOrFindParagraphStyleNode(DOMNode $parent)
    {
        $styleNode = null;
        foreach ($parent->childNodes as $childNode) {
            if ($childNode->nodeName === 'w:pPr') {
                return $childNode;
            }
        }
        if ($styleNode === null) {
            $fragment = $parent->ownerDocument->createDocumentFragment();
            $fragment->appendXML('<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"> </w:pPr>');
            $parent->insertBefore($fragment);
            return $this->addOrFindParagraphStyleNode($parent);
        }
        return $styleNode;
    }

    private function addParagraphDirection(DOMNode $parent)
    {
        $fragment = $parent->ownerDocument->createDocumentFragment();
        $direction = null;
        if ($this->direction === 'ltr') {
            $direction = 'start';
        } elseif ($this->direction === 'rtl') {
            $direction = 'end';
        }
        if ($direction !== null) {
            foreach ($parent->childNodes as $childNode) {
                if ($childNode->nodeName === 'w:jc') {
                    $parent->removeChild($childNode);
                }
            }
            $fragment->appendXML('<w:jc w:val="' . $direction . '" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />');
            $parent->appendChild($fragment);
        }
    }
}