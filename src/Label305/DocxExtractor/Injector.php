<?php namespace Label305\DocxExtractor;


use Label305\DocxExtractor\Decorated\Paragraph;

interface Injector {

    /**
     * @param Paragraph[]|array $mapping
     * @param string $fileToInjectLocationPath
     * @param string $saveLocationPath
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return void
     */
    public function injectMappingAndCreateNewFile(array $mapping, string $fileToInjectLocationPath, string $saveLocationPath):void;

}