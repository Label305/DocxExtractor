<?php

namespace Label305\DocxExtractor\Decorated;


use DOMElement;
use DOMNode;
use DOMText;
use Label305\DocxExtractor\DocxFileException;
use Label305\DocxExtractor\DocxHandler;
use Label305\DocxExtractor\DocxParsingException;
use Label305\DocxExtractor\Extractor;

class DecoratedTextExtractor extends DocxHandler implements Extractor
{

    /**
     * @var int
     */
    protected $nextTagIdentifier;

    /**
     * @param $originalFilePath
     * @param $mappingFileSaveLocationPath
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return array The mapping of all the strings
     */
    public function extractStringsAndCreateMappingFile($originalFilePath, $mappingFileSaveLocationPath)
    {
        $prepared = $this->prepareDocumentForReading($originalFilePath);

        $this->nextTagIdentifier = 0;
        $result = $this->replaceAndMapValues($prepared['dom']->documentElement);
        $this->saveDocument($prepared['dom'], $prepared["archive"], $mappingFileSaveLocationPath);

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

        if ($node instanceof DOMElement && $node->nodeName == "w:p") {
            $result = array_merge(
                $result,
                $this->replaceAndMapValuesForParagraph($node)
            );
        } else {
            if ($node->childNodes !== null) {
                foreach ($node->childNodes as $child) {
                    $result = array_merge(
                        $result,
                        $this->replaceAndMapValues($child)
                    );
                }
            }
        }

        return $result;
    }

    /**
     * @param DOMNode $paragraph
     * @return array
     */
    protected function replaceAndMapValuesForParagraph(DOMNode $paragraph)
    {
        $result = [];

        if ($paragraph->childNodes !== null) {

            $firstTextChild = null;
            $otherNodes = [];
            $parts = new Paragraph();

            $nodeNames = [
                "w:r",
                "w:hyperlink",
                "w:smartTag",
                "mc:AlternateContent"
            ];

            foreach ($paragraph->childNodes as $paragraphChild) {
                if ($paragraphChild instanceof DOMElement && in_array($paragraphChild->nodeName, $nodeNames)) {
                    $paragraphParts = $this->parseRNode($paragraphChild);

                    if (count($paragraphParts) !== 0) {
                        foreach ($paragraphParts as $paragraphPart) {
                            $parts[] = $paragraphPart;
                        }

                        if ($firstTextChild === null) {
                            $firstTextChild = $paragraphChild;
                        } else {
                            $otherNodes[] = $paragraphChild;
                        }
                    }
                }
            }

            if ($firstTextChild !== null) {
                $replacementNode = new DOMText();
                $replacementNode->nodeValue = "%" . $this->nextTagIdentifier . "%";
                $paragraph->replaceChild($replacementNode, $firstTextChild);

                foreach ($otherNodes as $otherNode) {
                    $paragraph->removeChild($otherNode);
                }

                $result[$this->nextTagIdentifier] = $parts;
                $this->nextTagIdentifier++;
            }
        }

        return $result;
    }

    /**
     * @param DOMElement $rNode
     * @return array
     */
    protected function parseRNode(DOMElement $rNode)
    {
        $bold = false;
        $italic = false;
        $underline = false;
        $brCount = 0;
        $highLight = false;
        $superscript = false;
        $subscript = false;
        $text = null;
        $result = [];

        foreach ($rNode->childNodes as $rChild) {
            $this->parseChildNode($rChild, $result, $bold, $italic, $underline, $brCount, $highLight,
                $superscript, $subscript, $text
            );
        }

        return $result;
    }

    /**
     * @param $rChild
     * @param $result
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
        &$bold,
        &$italic,
        &$underline,
        &$brCount,
        &$highLight,
        &$superscript,
        &$subscript,
        &$text
    ) {
        if ($rChild instanceof DOMElement && in_array($rChild->nodeName, ["mc:AlternateContent", "mc:Fallback"])) {
            foreach ($rChild->childNodes as $childNode) {
                if ($childNode->nodeName === 'mc:Fallback') {
                    $text = trim(implode($this->parseText($childNode)));
                } else {
                    $this->parseChildNode($childNode, $result, $bold, $italic, $underline, $brCount, $highLight,
                        $superscript, $subscript, $text);
                }
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
                    $this->parseChildNode($propertyNode, $result, $bold, $italic, $underline, $brCount, $highLight,
                        $superscript, $subscript, $text);
                }
            }
        } elseif ($rChild instanceof DOMElement && $rChild->nodeName == "w:rPr") {
            foreach ($rChild->childNodes as $propertyNode) {
                if ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:b") {
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
                    }
                    if ($variant === 'subscript') {
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

        if ($text !== null) {
            $result[] = new Sentence($text, $bold, $italic, $underline, $brCount, $highLight, $superscript, $subscript);
            $brCount = 0;
            $text = null;
        } elseif ($brCount !== 0) {
            $result[] = new Sentence('', $bold, $italic, $underline, $brCount, $highLight, $superscript, $subscript);
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