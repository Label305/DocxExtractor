<?php

namespace Label305\DocxExtractor\Decorated;


use DOMElement;
use DOMNode;
use DOMText;
use Label305\DocxExtractor\DocxFileException;
use Label305\DocxExtractor\DocxHandler;
use Label305\DocxExtractor\DocxParsingException;
use Label305\DocxExtractor\Extractor;

class DecoratedTextExtractor extends DocxHandler implements Extractor {

    /**
     * @var int
     */
    protected $nextTagIdentifier;

    /**
     * @param $originalFilePath
     * @param $mappingFileSaveLocationPath
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return Array The mapping of all the strings
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
            ];

            foreach ($paragraph->childNodes as $paragraphChild) {
                if ($paragraphChild instanceof DOMElement && in_array($paragraphChild->nodeName, $nodeNames)) {

                    $paragraphPart = $this->parseRNode($paragraphChild);
                    if ($paragraphPart !== null) {
                        $parts[] = $paragraphPart;
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

    protected function parseRNode(DOMElement $rNode)
    {
        $bold = false;
        $italic = false;
        $underline = false;
        $brCount = 0;
        $text = null;

        foreach ($rNode->childNodes as $rChild) {

            if ($rChild instanceof DOMElement && in_array($rChild->nodeName, ["w:r", "w:smartTag"])) {
                foreach ($rChild->childNodes as $propertyNode) {
                    if ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:t") {
                        $text = trim(implode($this->parseText($rChild)), " ");
                    } else {
                        $text = trim(implode($this->parseText($rChild)), " ");
                    }
                }
            }
            
            elseif ($rChild instanceof DOMElement && $rChild->nodeName == "w:rPr") {
                foreach ($rChild->childNodes as $propertyNode) {
                    if ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:b") {
                        $bold = true;
                    }
                    elseif ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:i") {
                        $italic = true;
                    }
                    elseif ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:u") {
                        $underline = true;
                    }
                }
            }

            elseif ($rChild instanceof DOMElement && $rChild->nodeName == "w:t") {
                if ($rChild->getAttribute("xml:space") == 'preserve') {
                    $text = implode($this->parseText($rChild));
                } else {
                    $text = trim(implode($this->parseText($rChild)), " ");
                }

            }

            elseif ($rChild instanceof DOMElement && $rChild->nodeName == "w:br") {
                $brCount++;
            }
        }

        if ($text != null) {
            return new Sentence($text, $bold, $italic, $underline, $brCount);
        } else {
            return null;
        }
    }

    protected function parseText(DOMNode $node)
    {
        $result = [];

        if ($node instanceof DOMText) {
            $result[] = $node->nodeValue;
        }

        if ($node->childNodes !== null) {
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