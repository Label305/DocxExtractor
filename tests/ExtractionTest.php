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

        $this->assertEquals("Haynes-Shockley experiment", $mapping[0]);
        $this->assertEquals("Practical work P3, solid state physics", $mapping[1]);
        $this->assertEquals("Introduction", $mapping[2]);

        $mapping[0] = "Haynes-Shockley experiment updated";
        $mapping[1] = "Practical work P3, solid state physics updated";
        $mapping[2] = "Introduction updated";

        $injector = new BasicInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/crazy-extracted.docx', __DIR__.'/fixtures/crazy-injected.docx');

        $otherExtractor = new BasicExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/crazy-injected.docx', __DIR__.'/fixtures/crazy-injected-extracted.docx');

        $this->assertEquals("Haynes-Shockley experiment updated", $otherMapping[0]);
        $this->assertEquals("Practical work P3, solid state physics updated", $otherMapping[1]);
        $this->assertEquals("Introduction updated", $otherMapping[2]);

        unlink(__DIR__.'/fixtures/crazy-extracted.docx');
        unlink(__DIR__.'/fixtures/crazy-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/crazy-injected.docx');
    }

    public function testTagMappingDecoratedExtractorWithSimpleDocument() {

        $extractor = new DecoratedTextExtractor();

        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/simple.docx', __DIR__.'/fixtures/simple-extracted.docx');

        $this->assertEquals("The quick brown fox jumps over the lazy dog", $mapping[0][0]->text);

        $mapping[0][0]->text = "Several fabulous dixieland jazz groups played with quick tempo.";

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

    public function testTagMappingDecoratedExtractorWithDocumentContainingHyperlink() {

        $extractor = new DecoratedTextExtractor();

        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/hyperlink.docx', __DIR__.'/fixtures/hyperlink-extracted.docx');

        $this->assertEquals("Bent u geïnteresseerd in een nieuw gebouwde ruime woning vanaf Euro ", $mapping[0][0]->text);
        $this->assertEquals("69.000,–", $mapping[0][1]->text);
        $this->assertEquals("? ", $mapping[0][2]->text);
        $this->assertEquals("KLIK OP DEZE LINK EN ZIE UW NIEUW GEBOUWDE WONING.", $mapping[0][3]->text);

        $mapping[0][0]->text = Paragraph::paragraphWithHTML("Are you interested in a newly built spacious house from Euro&nbsp;")->toHTML();
        $mapping[0][1]->text = Paragraph::paragraphWithHTML("69.000,&ndash;")->toHTML();
        $mapping[0][2]->text = "? ";
        $mapping[0][3]->text = Paragraph::paragraphWithHTML("CLICK ON THIS LINK AND SEE YOUR NEW BUILD HOUSE.")->toHTML();

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/hyperlink-extracted.docx', __DIR__.'/fixtures/hyperlink-injected.docx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/hyperlink-injected.docx', __DIR__.'/fixtures/hyperlink-injected-extracted.docx');

        $this->assertEquals("Are you interested in a newly built spacious house from Euro ", $otherMapping[0][0]->text);
        $this->assertEquals("69.000,&ndash;", $otherMapping[0][1]->text);
        $this->assertEquals("? ", $otherMapping[0][2]->text);
        $this->assertEquals("CLICK ON THIS LINK AND SEE YOUR NEW BUILD HOUSE.", $otherMapping[0][3]->text);

        unlink(__DIR__.'/fixtures/hyperlink-extracted.docx');
        unlink(__DIR__.'/fixtures/hyperlink-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/hyperlink-injected.docx');
    }

    public function testTagMappingDecoratedExtractorWithDocumentContainingSmartTag() {

        $extractor = new DecoratedTextExtractor();

        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/smart_tag.docx', __DIR__.'/fixtures/smart_tag-extracted.docx');

        // There are smart tags
        $this->assertEquals("France", $mapping[5][3]->text);
        $this->assertEquals("Gascony", $mapping[26][3]->text);
        $this->assertEquals("25 km", $mapping[31][3]->text);
        $this->assertEquals("Tarbes", $mapping[35][4]->text);
        $this->assertEquals("Tarbes", $mapping[35][6]->text);
        $this->assertEquals("Achater", $mapping[50][5]->text);

        $mapping[5][3]->text = Paragraph::paragraphWithHTML("Frankrijk")->toHTML();
        $mapping[26][3]->text = Paragraph::paragraphWithHTML("Gascony vertaald")->toHTML();
        $mapping[31][3]->text = Paragraph::paragraphWithHTML("25 kilometer")->toHTML();
        $mapping[35][4]->text = Paragraph::paragraphWithHTML("Tarbes vertaald")->toHTML();
        $mapping[35][6]->text = Paragraph::paragraphWithHTML("Tarbes vertaald 2")->toHTML();
        $mapping[50][5]->text = Paragraph::paragraphWithHTML("Achater vertaald")->toHTML();

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/smart_tag-extracted.docx', __DIR__.'/fixtures/smart_tag-injected.docx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/smart_tag-injected.docx', __DIR__.'/fixtures/smart_tag-injected-extracted.docx');

        $this->assertEquals("Frankrijk", $otherMapping[5][3]->text);
        $this->assertEquals("Gascony vertaald", $otherMapping[26][3]->text);
        $this->assertEquals("25 kilometer", $otherMapping[31][3]->text);
        $this->assertEquals("Tarbes vertaald", $otherMapping[35][4]->text);
        $this->assertEquals("Tarbes vertaald 2", $otherMapping[35][6]->text);
        $this->assertEquals("Achater vertaald", $otherMapping[50][5]->text);

        unlink(__DIR__.'/fixtures/smart_tag-extracted.docx');
        unlink(__DIR__.'/fixtures/smart_tag-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/smart_tag-injected.docx');
    }

    public function testHyperlinks2InDocument() {

        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/hyperlinks_2.docx', __DIR__.'/fixtures/hyperlinks_2-extracted.docx');

        $this->assertEquals("Meer weten over deze doopsuikerdoosjes", $mapping[9][0]->text); // This is a link

        $mapping[9][0]->text = "Link vertaald";

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/hyperlinks_2-extracted.docx', __DIR__.'/fixtures/hyperlinks_2-injected.docx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/hyperlinks_2-injected.docx', __DIR__.'/fixtures/hyperlinks_2-injected-extracted.docx');

        $this->assertEquals("Link vertaald", $otherMapping[9][0]->text);

        unlink(__DIR__.'/fixtures/hyperlinks_2-extracted.docx');
        unlink(__DIR__.'/fixtures/hyperlinks_2-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/hyperlinks_2-injected.docx');
    }

    public function testHyperlinksInTextInDocument() {

        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/hyperlinks_in_text.docx', __DIR__.'/fixtures/hyperlinks_in_text-extracted.docx');

        $this->assertEquals("De website ", $mapping[1][0]->text);
        $this->assertEquals("www.apcoa.nl", $mapping[1][1]->text);
        $this->assertEquals(" en de daaraan gekoppelde internetdiensten worden u ter beschikking gesteld door APCOA PARKING Nederland B.V. (hierna: APCOA PARKING), statutair gevestigd te Rotterdam en ingeschreven in het handelsregister van de Kamer van Koophandel Provincie onder nummer ", $mapping[1][2]->text);

        $mapping[1][0]->text = "The website ";
        $mapping[1][1]->text = "www.apcoa.nl";
        $mapping[1][2]->text = " and the associated internet services are made available to you by APCOA PARKING Nederland B.V. (hereinafter: APCOA PARKING)";

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/hyperlinks_in_text-extracted.docx', __DIR__.'/fixtures/hyperlinks_in_text-injected.docx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/hyperlinks_in_text-injected.docx', __DIR__.'/fixtures/hyperlinks_in_text-injected-extracted.docx');

        $this->assertEquals("The website ", $otherMapping[1][0]->text);
        $this->assertEquals("www.apcoa.nl", $otherMapping[1][1]->text);
        $this->assertEquals(" and the associated internet services are made available to you by APCOA PARKING Nederland B.V. (hereinafter: APCOA PARKING)", $otherMapping[1][2]->text);

        unlink(__DIR__.'/fixtures/hyperlinks_in_text-extracted.docx');
        unlink(__DIR__.'/fixtures/hyperlinks_in_text-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/hyperlinks_in_text-injected.docx');
    }

    public function testMarkingsInDocument() {

        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/markings.docx', __DIR__.'/fixtures/markings-extracted.docx');

        $this->assertEquals("Energie gebruik ", $mapping[53][2]->text); // This is highlight
        $this->assertEquals("volgens leveranciers opgave ", $mapping[53][3]->text); // This is highlight
        $this->assertEquals("van A-merk", $mapping[53][4]->text); // This is highlight

        $mapping[53][2]->text = "Dit is ";
        $mapping[53][3]->text = "het vertaalde ";
        $mapping[53][4]->text = "stuk tekst";

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__.'/fixtures/markings-extracted.docx', __DIR__.'/fixtures/markings-injected.docx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/markings-injected.docx', __DIR__.'/fixtures/markings-injected-extracted.docx');

        $this->assertEquals("Dit is ", $otherMapping[53][2]->text);
        $this->assertEquals("het vertaalde ", $otherMapping[53][3]->text);
        $this->assertEquals("stuk tekst", $otherMapping[53][4]->text);

        unlink(__DIR__.'/fixtures/markings-extracted.docx');
        unlink(__DIR__.'/fixtures/markings-injected-extracted.docx');
        unlink(__DIR__.'/fixtures/markings-injected.docx');
    }

    public function testGetSuperscriptInDocument() {

        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/superscript.docx', __DIR__.'/fixtures/superscript-extracted.docx');

        $this->assertEquals("1", $mapping[0][3]->text);
        $this->assertTrue($mapping[0][3]->superscript);
        $this->assertFalse($mapping[0][3]->subscript);

        unlink(__DIR__.'/fixtures/superscript-extracted.docx');
    }

    public function testGetSubscriptInDocument() {

        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/subscript.docx', __DIR__.'/fixtures/subscript-extracted.docx');

        $this->assertEquals("1", $mapping[0][1]->text);
        $this->assertTrue($mapping[0][1]->subscript);
        $this->assertFalse($mapping[0][1]->superscript);

        unlink(__DIR__.'/fixtures/subscript-extracted.docx');
    }

}