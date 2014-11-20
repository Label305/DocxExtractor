<?php namespace Label305\DocxExtractor;


interface Extractor {

    /**
     * @param $originalFilePath
     * @param $mappingFileSaveLocationPath
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return Array The mapping of all the strings
     */
    public function extractStringsAndCreateMappingFile($originalFilePath, $mappingFileSaveLocationPath);

}