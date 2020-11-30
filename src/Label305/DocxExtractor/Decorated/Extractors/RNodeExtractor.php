<?php

namespace Label305\DocxExtractor\Decorated\Extractors;

use DOMElement;
use DOMNode;
use DOMText;
use Label305\DocxExtractor\Decorated\Sentence;
use Label305\DocxExtractor\Decorated\Style;

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
        $style = null;
        $result = [];

        foreach ($DOMElement->childNodes as $rChild) {
            $this->parseChildNode($rChild, $result, $webHidden, $bold, $italic, $underline, $brCount, $highLight,
                $superscript, $subscript, $text, $style
            );
        }

        return $result;
    }

    /**
     * @param $rChild
     * @param array $result
     * @param bool $webHidden
     * @param bool $bold
     * @param bool $italic
     * @param bool $underline
     * @param int $brCount
     * @param bool $highLight
     * @param bool $superscript
     * @param bool $subscript
     * @param string|null $text
     * @param Style|null $style
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
        &$text,
        &$style
    ) {
        if ($rChild instanceof DOMElement) {
            switch ($rChild->nodeName) {
                case "w:p" :
                case "w:r" :
                case "w:smartTag" :
                    foreach ($rChild->childNodes as $propertyNode) {
                        $this->parseChildNode($propertyNode, $result, $webHidden, $bold, $italic, $underline, $brCount, $highLight,
                            $superscript, $subscript, $text, $style);
                    }
                    break;

                case "w:rPr" :
                    $rFonts = null;
                    $color = null;
                    $lang = null;
                    $sz = null;
                    $szCs = null;
                    $position = null;
                    $spacing = null;
                    $highLightColor = null;
                    $hasStyle = false;

                    foreach ($rChild->childNodes as $propertyNode) {
                        if ($propertyNode instanceof DOMElement) {
                            $this->parseStyle($propertyNode,$rFonts,$color,$lang,$sz,$szCs, $position, $spacing, $highLightColor, $hasStyle);
                            $this->parseFormatting($propertyNode,$webHidden,$bold,$italic,$underline,$highLight,$superscript,$subscript);
                        }
                    }

                    if ($hasStyle) {
                        $style = new Style($rFonts, $color, $lang, $sz, $szCs, $position, $spacing, $highLightColor);
                    }
                    break;

                case "w:t" :
                case "w:tab" :
                    if ($rChild->nodeName == "w:tab") {
                        $text = " ";
                    } else {
                        $text = implode(" ", $this->parseText($rChild));
                    }
                    break;

                case "w:br" :
                    $brCount++;
                    break;
            }
        }

        if (!$webHidden && ($brCount !== 0 || ($text !== null && strlen($text) !== 0))) {

            $result[] = new Sentence($text, $bold, $italic, $underline, $brCount, $highLight, $superscript, $subscript, $style);

            // Reset
            $brCount = 0;
            $style = null;
            $text = null;
        }
    }

    /**
     * @param DOMElement $propertyNode
     * @param bool $webHidden
     * @param bool $bold
     * @param bool $italic
     * @param bool $underline
     * @param bool $highLight
     * @param bool $superscript
     * @param bool $subscript
     */
    protected function parseFormatting(
        DOMElement $propertyNode,
        &$webHidden,
        &$bold,
        &$italic,
        &$underline,
        &$highLight,
        &$superscript,
        &$subscript
    ) {
        if ($propertyNode->nodeName == "w:webHidden") {
            $webHidden = true;
        } elseif ($propertyNode->nodeName == "w:b") {
            $bold = true;
        } elseif ($propertyNode->nodeName == "w:i") {
            $italic = true;
        } elseif ($propertyNode->nodeName == "w:u") {
            $underline = true;
        } elseif ($propertyNode->nodeName == "w:highlight") {
            $highLight = true;
        } elseif ($propertyNode->nodeName == "w:vertAlign") {
            $variant = $propertyNode->getAttribute('w:val');
            if ($variant === 'superscript') {
                $superscript = true;
            } elseif ($variant === 'subscript') {
                $subscript = true;
            }
        }
    }

    /**
     * @param DOMElement $propertyNode
     * @param string|null $rFonts
     * @param string|null $color
     * @param string|null $lang
     * @param string|null $sz
     * @param string|null $szCs
     * @param string|null $position
     * @param string|null $spacing
     * @param string|null $highLightColor
     * @param bool $hasStyle
     */
    protected function parseStyle(DOMElement $propertyNode, &$rFonts, &$color, &$lang, &$sz, &$szCs, &$position, &$spacing, &$highLightColor, &$hasStyle)
    {
        if ($propertyNode->nodeName == "w:rFonts") {
            $rFonts = $propertyNode->getAttribute('w:ascii');
            if ($rFonts === null) {
                $rFonts = $propertyNode->getAttribute('w:hAnsi');
            }
            $hasStyle = true;
        } elseif($propertyNode->nodeName == "w:color") {
            $color = $propertyNode->getAttribute('w:val');
            $hasStyle = true;
        } elseif($propertyNode->nodeName == "w:lang") {
            $lang = $propertyNode->getAttribute('w:val');
            $hasStyle = true;
        } elseif($propertyNode->nodeName == "w:sz") {
            $sz = $propertyNode->getAttribute('w:val');
            $hasStyle = true;
        } elseif($propertyNode->nodeName == "w:szCs") {
            $szCs = $propertyNode->getAttribute('w:val');
            $hasStyle = true;
        } elseif ($propertyNode->nodeName == "w:position") {
            $position = $propertyNode->getAttribute('w:val');
            $hasStyle = true;
        } elseif ($propertyNode->nodeName == "w:spacing") {
            $spacing = $propertyNode->getAttribute('w:val');
            $hasStyle = true;
        } elseif ($propertyNode->nodeName == "w:highlight") {
            $highLightColor = $propertyNode->getAttribute('w:val');
            $hasStyle = true;
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