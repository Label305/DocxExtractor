<?php
/**
 * Created by PhpStorm.
 * User: Thijs
 * Date: 13-11-14
 * Time: 09:45
 */

namespace Label305\DocxExtractor;


use DOMElement;
use DOMNode;
use DOMText;

class DecoratedTextExtractor extends DocxHandler implements Extractor {

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
            $parts = [];

            foreach ($paragraph->childNodes as $paragraphChild) {
                if ($paragraphChild instanceof DOMElement && $paragraphChild->nodeName == "w:r") {
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
                $replacementNode->nodeValue = "%" . $this->nextTagIdentifier . "-p%";
                $paragraph->replaceChild($replacementNode, $firstTextChild);

                foreach ($otherNodes as $otherNode) {
                    $paragraph->removeChild($otherNode);
                }

                $result[$this->nextTagIdentifier . "-p"] = $parts;
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

            if ($rChild instanceof DOMElement && $rChild->nodeName == "w:rPr") {
                foreach ($rChild->childNodes as $propertyNode) {
                    if ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:b") {
                        $bold = true;
                    }
                    if ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:i") {
                        $italic = true;
                    }
                    if ($propertyNode instanceof DOMElement && $propertyNode->nodeName == "w:u") {
                        $underline = true;
                    }
                }
            }

            if ($rChild instanceof DOMElement && $rChild->nodeName == "w:t") {
                $text = implode($this->parseText($rChild));
            }

            if ($rChild instanceof DOMElement && $rChild->nodeName == "w:br") {
                $brCount++;
            }
        }

        if ($text != null) {
            return [
                "text" => $text,
                "bold" => $bold,
                "italic" => $italic,
                "underline" => $underline,
                "br_count" => $brCount
            ];
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