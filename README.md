Docx Extractor [![Build Status](https://travis-ci.org/Label305/DocxExtractor.svg)](https://travis-ci.org/Label305/DocxExtractor)
=============

PHP library for extracting and replacing string data in .docx files. Docx files are zip archives filled with XML documents and assets. Their format is described by [OOXML](http://nl.wikipedia.org/wiki/Office_Open_XML). This library only manipulates the `word/document.xml` file.

Composer installation
---

```json
"require": {
    "label305/docx-extractor": "0.1.*"
}
```

Basic usage
----

Import the basic classes.

```php
use Label305\DocxExtractor\Basic\BasicExtractor;
use Label305\DocxExtractor\Basic\BasicInjector;
```

First we need to extract all the paragraphs from an existing `docx` file. This can be done using the `BasicExtractor` or the `DecoratedTextExtractor`. Calling `extractStringsAndCreateMappingFile` will create a new file which name you pass in the second argument. This new file contains references so the library knows where to later inject the altered text back into.

```php
$extractor = new BasicExtractor();
$mapping = $extractor->extractStringsAndCreateMappingFile(
    'simple.docx',
    'simple-extracted.docx'
  );
```

Now that you have extracted paragraphs you can inspect the content of the resulting `$mapping` array. And if you wish to change the content you can simply modify it. The array key maps to a symbol in the `simple-extracted.docx`.

```php
echo $mapping[0]; // The quick brown fox jumps over the lazy dog
```

Now after you changed your content, you can save it back to a new file. In this case that file is `simple-injected.docx`.

```php
$mapping[0] = "Several fabulous dixieland jazz groups played with quick tempo.";

$injector = new BasicInjector();
$injector->injectMappingAndCreateNewFile(
    $mapping,
    'simple-extracted.docx',
    'simple-injected.docx'
  );
```

Advanced usage
----

The library is also equiped with a `DecoratedTextExtractor` and `DecoratedTextInjector` with which you can manipulate basic paragraph styling like bold, italic and underline. You can also use the `Paragraph` objects to distinguish logical groupings of text.

```php
$extractor = new DecoratedTextExtractor();
$mapping = $extractor->extractStringsAndCreateMappingFile(
    'simple.docx',
    'simple-extracted.docx'
  );
  
$firstParagraph = $mapping[0]; // Paragraph object
$firstSentence = $firstParagraph[0]; // Sentence object

$firstSentence->italic = true;
$firstSentence->bold = false;
$firstSentence->underline = true;
$firstSentence->br = 2; // Two line breaks before this sentence

echo $firstSentence->text; // The quick brown fox jumps over the lazy dog
$firstSentence->text = "Several fabulous dixieland jazz groups played with quick tempo.";

$injector = new DecoratedTextInjector();
$injector->injectMappingAndCreateNewFile(
    $mapping,
    'simple-extracted.docx',
    'simple-injected.docx'
  );
```

License
---------
Copyright 2014 Label305 B.V.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

[http://www.apache.org/licenses/LICENSE-2.0](http://www.apache.org/licenses/LICENSE-2.0)

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
