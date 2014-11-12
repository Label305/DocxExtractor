<?php namespace Label305\DocxExtractor;


interface Injector {

    /**
     * @param $mapping
     * @param $fileToInjectLocationHandle
     * @param $saveLocationHandle
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return void
     */
    public function injectMappingAndCreateNewFile($mapping, $fileToInjectLocationHandle, $saveLocationHandle);

}