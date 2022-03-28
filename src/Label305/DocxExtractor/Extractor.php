<?php namespace Label305\DocxExtractor;


interface Extractor {

    /**
     * @param string $originalFilePath
     * @param string $mappingFileSaveLocationPath
     * @throws DocxParsingException
     * @throws DocxFileException
     * @return array The mapping of all the strings
     */
    public function extractStringsAndCreateMappingFile(string $originalFilePath, string $mappingFileSaveLocationPath): array;

}