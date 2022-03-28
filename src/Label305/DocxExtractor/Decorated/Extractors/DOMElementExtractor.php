<?php

namespace Label305\DocxExtractor\Decorated\Extractors;

use DOMElement;

interface DOMElementExtractor
{
    /**
     * @param DOMElement $DOMElement
     * @return DOMElement|null
     */
    public function extract(DOMElement $DOMElement): ?DOMElement;
}