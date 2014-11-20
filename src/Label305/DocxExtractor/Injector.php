<?php namespace Label305\DocxExtractor;


interface Injector {

    /**
     * @param $mapping
     * @param $fileToInjectLocationPath
     * @param $saveLocationPath
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return void
     */
    public function injectMappingAndCreateNewFile($mapping, $fileToInjectLocationPath, $saveLocationPath);

}