<?php

namespace Label305\DocxExtractor\Decorated\Extractors;

use DOMElement;

class TextBoxDOMElementExtractor implements DOMElementExtractor
{
    /**
     * @param DOMElement $DOMElement
     * The result is the array which contains te sentences
     * @return DOMElement
     */
    public function extract(DOMElement $DOMElement)
    {
        if ($DOMElement->nodeName == "mc:AlternateContent") {
            foreach ($DOMElement->childNodes as $childChildNode) {
                // Only translate the primary content
                if ($childChildNode instanceof DOMElement && $childChildNode->nodeName == "mc:Choice") {
                    return $childChildNode;
                }
            }
        }
    }
}