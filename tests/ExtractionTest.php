<?php

use Label305\DocxExtractor\BasicExtractor;
use Label305\DocxExtractor\BasicInjector;
use Label305\DocxExtractor\DecoratedTextExtractor;

class ExtractionTest extends TestCase {

    public function testTagMappingBasicExtractorWithSimpleDocument() {

        $extractor = new BasicExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/simple.docx', __DIR__.'/fixtures/simple-extracted.docx');

        $this->assertEquals("The quick brown fox jumps over the lazy dog", $mapping[0]);

        $mapping[0] = "Several fabulous dixieland jazz groups played with quick tempo.";

        $injector = new BasicInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/simple-extracted.docx', __DIR__.'/fixtures/simple-injected.docx');

        $otherExtractor = new BasicExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/simple-injected.docx', __DIR__.'/fixtures/simple-injected-extracted.docx');

        $this->assertEquals("Several fabulous dixieland jazz groups played with quick tempo.", $otherMapping[0]);

        unlink(__DIR__.'/fixtures/simple-extracted.docx');
        unlink(__DIR__.'/fixtures/simple-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/simple-injected.docx');
    }

    public function testTagMappingBasicExtractorWithCrazyDocument() {

        $extractor = new BasicExtractor();

        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/crazy.docx', __DIR__.'/fixtures/crazy-extracted.docx');
        //var_dump($mapping);

        unlink(__DIR__.'/fixtures/crazy-extracted.docx');
    }

    public function testTagMappingDecoratedExtractorWithSimpleDocument() {

        $extractor = new DecoratedTextExtractor();

        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/simple.docx', __DIR__.'/fixtures/simple-extracted.docx');
        //var_dump($mapping);

        $this->assertEquals("The quick brown fox jumps over the lazy dog", $mapping["0-p"][0]["text"]);

        unlink(__DIR__.'/fixtures/simple-extracted.docx');
    }

}