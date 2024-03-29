<?php

namespace Label305\DocxExtractor\Decorated;

use DOMElement;
use DOMNode;
use DOMText;
use Label305\DocxExtractor\Decorated\Extractors\RNodeSentenceExtractor;
use Label305\DocxExtractor\Decorated\Extractors\TextBoxDOMElementExtractor;
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
     * @param string $originalFilePath
     * @param string $mappingFileSaveLocationPath
     * @throws DocxParsingException
     * @throws DocxFileException
     * The result is the mapping of all the strings
     * @return Paragraph[]|array
     */
    public function extractStringsAndCreateMappingFile(string $originalFilePath, string $mappingFileSaveLocationPath): array
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
    protected function replaceAndMapValues(DOMNode $node): array
    {
        $result = [];

        if ($node instanceof DOMElement && $node->nodeName === "w:p") {
            $this->replaceAndMapValuesForParagraph($node, $result);
        }
        else {
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
     * @param DOMNode $DOMNode
     * @param array $result
     * @return array
     */
    protected function replaceAndMapValuesForParagraph(DOMNode $DOMNode, array &$result): array
    {
        if ($DOMNode->childNodes !== null) {

            $firstTextChild = null;
            $otherNodes = [];
            $parts = new Paragraph();

            foreach ($DOMNode->childNodes as $DOMNodeChild) {
                if ($DOMNodeChild instanceof DOMElement && in_array($DOMNodeChild->nodeName, [
                    "w:r",
                    "w:ins",
                    "w:del",
                    "w:hyperlink",
                    "w:smartTag",
                ])) {
                    // Additional loops for specific elements
                    if ($DOMNodeChild->childNodes !== null) {
                        foreach ($DOMNodeChild->childNodes as $childNode) {
                            switch ($childNode->nodeName) {
                                case "mc:AlternateContent" :
                                    $textBoxChild = (new TextBoxDOMElementExtractor())->extract($childNode);
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
                    $sentences = (new RNodeSentenceExtractor())->extract($DOMNodeChild);
                    if (count($sentences) !== 0) {
                        foreach ($sentences as $sentence) {
                            $parts[] = $sentence;
                        }
                        if ($firstTextChild === null) {
                            $firstTextChild = $DOMNodeChild;
                        } else {
                            $otherNodes[] = $DOMNodeChild;
                        }
                    }

                } elseif ($DOMNodeChild instanceof DOMElement) {
                    if ($DOMNodeChild->nodeName === "w:sdtContent") {
                        $this->replaceAndMapValues($DOMNodeChild);
                    } else {
                        $this->replaceAndMapValuesForParagraph($DOMNodeChild, $result);
                    }
                }
            }

            if ($firstTextChild !== null) {
                $replacementNode = new DOMText();
                $replacementNode->nodeValue = "%" . $this->nextTagIdentifier . "%";
                $DOMNode->replaceChild($replacementNode, $firstTextChild);

                foreach ($otherNodes as $otherNode) {
                    $DOMNode->removeChild($otherNode);
                }

                $result[$this->nextTagIdentifier] = $parts;
                $this->nextTagIdentifier++;
            }
        }

        return $result;
    }
}