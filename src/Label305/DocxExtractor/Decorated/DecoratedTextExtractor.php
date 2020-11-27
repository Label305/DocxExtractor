<?php

namespace Label305\DocxExtractor\Decorated;

use DOMElement;
use DOMNode;
use DOMText;
use Label305\DocxExtractor\Decorated\Extractors\RNodeExtractor;
use Label305\DocxExtractor\Decorated\Extractors\TextBoxExtractor;
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
            $this->replaceAndMapValuesForParagraph($node, $result);

        } elseif ($node instanceof DOMElement && $node->nodeName == "w:sdtContent") {
            if ($node->childNodes !== null) {
                foreach ($node->childNodes as $child) {
                    if ($child instanceof DOMElement && $child->nodeName == "w:p") {
                        $this->replaceAndMapValuesForParagraph($node, $result);
                    }
                }
            }

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
     * @param $result
     * @return array
     */
    protected function replaceAndMapValuesForParagraph(DOMNode $paragraph, &$result)
    {
        if ($paragraph->childNodes !== null) {

            $firstTextChild = null;
            $otherNodes = [];
            $parts = new Paragraph();

            foreach ($paragraph->childNodes as $paragraphChild) {
                if ($paragraphChild instanceof DOMElement && in_array($paragraphChild->nodeName, [
                    "w:r",
                    "w:hyperlink",
                    "w:smartTag",
                ])) {

                    // Additional loops for specific elements
                    if ($paragraphChild->childNodes !== null) {
                        foreach ($paragraphChild->childNodes as $childNode) {
                            switch ($childNode->nodeName) {
                                case "mc:AlternateContent" :
                                    $textBoxChild = (new TextBoxExtractor())->extract($childNode);
                                    if ($textBoxChild !== null) {
                                        $this->replaceAndMapValuesForParagraph($textBoxChild, $result);
                                    }
                                    break;
                                default :
                                    break;
                            }
                        }
                    }

                    // Parse results
                    $paragraphParts = (new RNodeExtractor())->extract($paragraphChild);
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

                } elseif ($paragraphChild instanceof DOMElement) {
                    $this->replaceAndMapValuesForParagraph($paragraphChild, $result);
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
}