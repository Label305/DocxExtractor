<?php

namespace Label305\DocxExtractor\Decorated\Extractors;

use DOMElement;
use DOMNode;
use DOMText;
use Label305\DocxExtractor\Decorated\Sentence;

class RNodeExtractor implements DecoratedExtractor
{
    /**
     * @param DOMElement $DOMElement
     * The result is the array which contains te sentences
     * @return Sentence[]
     */
    public function extract(DOMElement $DOMElement)
    {
        $webHidden = false;
        $bold = false;
        $italic = false;
        $underline = false;
        $brCount = 0;
        $highLight = false;
        $superscript = false;
        $subscript = false;
        $text = null;
        $result = [];

        foreach ($DOMElement->childNodes as $rChild) {
            $this->parseChildNode($rChild, $result, $webHidden, $bold, $italic, $underline, $brCount, $highLight,
                $superscript, $subscript, $text
            );
        }

        return $result;
    }

    /**
     * @param $rChild
     * @param $result
     * @param $webHidden
     * @param $bold
     * @param $italic
     * @param $underline
     * @param $brCount
     * @param $highLight
     * @param $superscript
     * @param $subscript
     * @param $text
     */
    private function parseChildNode(
        $rChild,
        &$result,
        &$webHidden,
        &$bold,
        &$italic,
        &$underline,
        &$brCount,
        &$highLight,
        &$superscript,
        &$subscript,
        &$text
    ) {
        if ($rChild instanceof DOMElement && in_array($rChild->nodeName, ["w:p"])) {
            foreach ($rChild->childNodes as $propertyNode) {
                $this->parseChildNode($propertyNode, $result, $webHidden, $bold, $italic, $underline, $brCount, $highLight,
                    $superscript, $subscript, $text);
            }

        } elseif ($rChild instanceof DOMElement && in_array($rChild->nodeName, ["w:r", "w:smartTag"])) {
            foreach ($rChild->childNodes as $propertyNode) {
                if ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:t") {
                    if ($propertyNode->getAttribute("xml:space") == 'preserve') {
                        $text = implode($this->parseText($propertyNode));
                    } else {
                        $text = trim(implode($this->parseText($propertyNode)), " ");
                    }
                } else {
                    $this->parseChildNode($propertyNode, $result, $webHidden, $bold, $italic, $underline, $brCount, $highLight,
                        $superscript, $subscript, $text);
                }
            }

        } elseif ($rChild instanceof DOMElement && $rChild->nodeName == "w:rPr") {
            foreach ($rChild->childNodes as $propertyNode) {
                if ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:webHidden") {
                    $webHidden = true;
                } elseif ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:b") {
                    $bold = true;
                } elseif ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:i") {
                    $italic = true;
                } elseif ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:u") {
                    $underline = true;
                } elseif ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:highlight") {
                    $highLight = true;
                } elseif ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:vertAlign") {
                    $variant = $propertyNode->getAttribute('w:val');
                    if ($variant === 'superscript') {
                        $superscript = true;
                    } elseif ($variant === 'subscript') {
                        $subscript = true;
                    }
                }
            }

        } elseif ($rChild instanceof DOMElement && $rChild->nodeName == "w:t") {
            if ($rChild->getAttribute("xml:space") == 'preserve') {
                $text = implode($this->parseText($rChild));
            } else {
                $text = trim(implode($this->parseText($rChild)), " ");
            }

        } elseif ($rChild instanceof DOMElement && $rChild->nodeName == "w:br") {
            $brCount++;
        }

        if (!$webHidden && ($brCount !== 0 || ($text !== null && strlen($text) !== 0))) {

            $result[] = new Sentence($text, $bold, $italic, $underline, $brCount, $highLight, $superscript, $subscript);
            $brCount = 0;
            $text = null;
        }
    }

    /**
     * @param DOMNode $node
     * @return array
     */
    protected function parseText(DOMNode $node)
    {
        $result = [];

        if ($node instanceof DOMText) {
            $result[] = $node->nodeValue;
        }

        if ($node->hasChildNodes()) {
            foreach ($node->childNodes as $child) {
                $result = array_merge(
                    $result,
                    $this->parseText($child)
                );
            }
        }
        return $result;
    }
}