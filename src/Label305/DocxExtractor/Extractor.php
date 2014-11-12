<?php namespace Label305\DocxExtractor;


interface Extractor {

    /**
     * @param $originalFileHandle
     * @param $mappingFileSaveLocationHandle
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return Array The mapping of all the strings
     */
    public function extractStringsAndCreateMappingFile($originalFileHandle, $mappingFileSaveLocationHandle);

}