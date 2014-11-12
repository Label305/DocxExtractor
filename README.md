DocxExtractor [![Build Status](https://travis-ci.org/Label305/DocxExtractor.svg)](https://travis-ci.org/Label305/DocxExtractor)
=============

PHP library for extracting and replacing string data in .docx files.

Basic usage
----

Import the basic classes.

```php
use Label305\DocxExtractor\BasicExtractor;
use Label305\DocxExtractor\BasicInjector;
```

And execute the extractor to get a mapping. Write your changes to the mapping and inject it back into the extracted file.

```php
$extractor = new BasicExtractor();
$mapping = $extractor->extractStringsAndCreateMappingFile(
    'simple.docx',
    'simple-extracted.docx'
  );

echo $mapping[0]; // The quick brown fox jumps over the lazy dog

$mapping[0] = "Several fabulous dixieland jazz groups played with quick tempo.";

$injector = new BasicInjector();
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
