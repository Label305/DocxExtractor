<?php

use Label305\DocxExtractor\Basic\BasicExtractor;
use Label305\DocxExtractor\Basic\BasicInjector;
use Label305\DocxExtractor\Decorated\DecoratedTextExtractor;
use Label305\DocxExtractor\Decorated\DecoratedTextInjector;
use Label305\DocxExtractor\Decorated\Paragraph;

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

        $this->assertEquals("The quick brown fox jumps over the lazy dog", $mapping[0][0]->text);

        $mapping[0][0]->text = "Several fabulous dixieland jazz groups played with quick tempo.";
        //var_dump($mapping[2]->toHTML());

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/simple-extracted.docx', __DIR__.'/fixtures/simple-injected.docx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/simple-injected.docx', __DIR__.'/fixtures/simple-injected-extracted.docx');

        $this->assertEquals("Several fabulous dixieland jazz groups played with quick tempo.", $otherMapping[0][0]->text);

        unlink(__DIR__.'/fixtures/simple-extracted.docx');
        unlink(__DIR__.'/fixtures/simple-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/simple-injected.docx');
    }

    public function testTagMappingDecoratedExtractorWithNormalDocument() {

        $extractor = new DecoratedTextExtractor();

        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/normal.docx', __DIR__.'/fixtures/normal-extracted.docx');

        $this->assertEquals("Aan", $mapping[0][0]->text);

        $mapping[0][0]->text = "At";
        $mapping[0][1]->text = "The";
        $mapping[0][2]->text = "Edge";
        //var_dump($mapping[2]->toHTML());

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/normal-extracted.docx', __DIR__.'/fixtures/normal-injected.docx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/normal-injected.docx', __DIR__.'/fixtures/normal-injected-extracted.docx');

        $this->assertEquals("At", $otherMapping[0][0]->text);

        unlink(__DIR__.'/fixtures/normal-extracted.docx');
        unlink(__DIR__.'/fixtures/normal-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/normal-injected.docx');
    }

    public function testTagMappingDecoratedExtractorWithCrazyDocument() {

        $extractor = new DecoratedTextExtractor();

        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/crazy.docx', __DIR__.'/fixtures/crazy-extracted.docx');
        //var_dump($mapping);

        unlink(__DIR__.'/fixtures/crazy-extracted.docx');
    }

    public function testTagMappingDecoratedExtractorWithNormalDocumentContainingNbspOrTilde() {

        $extractor = new DecoratedTextExtractor();

        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/normal.docx', __DIR__.'/fixtures/normal-extracted.docx');

        $this->assertEquals("Aan", $mapping[0][0]->text);

        $mapping[0][0]->text = Paragraph::paragraphWithHTML("At&nbsp;")->toHTML();
        $mapping[0][1]->text = Paragraph::paragraphWithHTML("The&nbsp;")->toHTML();
        $mapping[0][2]->text = "Edge";
        $mapping[0][3]->text = Paragraph::paragraphWithHTML("&Atilde;")->toHTML();

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/normal-extracted.docx', __DIR__.'/fixtures/normal-injected.docx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/normal-injected.docx', __DIR__.'/fixtures/normal-injected-extracted.docx');

        $this->assertEquals("At ", $otherMapping[0][0]->text);
        $this->assertEquals("The ", $otherMapping[0][1]->text);
        $this->assertEquals("Edge", $otherMapping[0][2]->text);
        $this->assertEquals("&Atilde;", $otherMapping[0][3]->text);

        unlink(__DIR__.'/fixtures/normal-extracted.docx');
        unlink(__DIR__.'/fixtures/normal-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/normal-injected.docx');
    }


}