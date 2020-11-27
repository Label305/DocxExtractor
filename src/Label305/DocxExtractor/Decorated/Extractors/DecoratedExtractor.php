<?php

namespace Label305\DocxExtractor\Decorated\Extractors;

use DOMElement;

interface DecoratedExtractor
{
    /**
     * @param DOMElement $DOMElement
     * @return mixed
     */
    public function extract(DOMElement $DOMElement);
}