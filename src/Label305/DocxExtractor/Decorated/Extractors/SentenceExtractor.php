<?php

namespace Label305\DocxExtractor\Decorated\Extractors;

use DOMElement;
use Label305\DocxExtractor\Decorated\Sentence;

interface SentenceExtractor
{
    /**
     * @param DOMElement $DOMElement
     * The result is the array which contains te sentences
     * @return Sentence[]
     */
    public function extract(DOMElement $DOMElement): array;
}