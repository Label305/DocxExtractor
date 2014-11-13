<?php namespace Label305\DocxExtractor;


use DOMDocument;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use ZipArchive;

abstract class DocxHandler {

    /**
     * @var string the tmp dir location
     */
    protected $temporaryDirectory = "/tmp";

    /**
     * Extract file
     * @param $fileHandle
     * @throws DocxFileException
     * @throws DocxParsingException
     * @returns array With "document" key, "dom" and "archive" key both are paths. "document" points to the document.xml
     * and "archive" points to the root of the archive. "dom" is the DOMDocument object for the document.xml.
     */
    protected function prepareDocumentForReading($fileHandle)
    {
        $temp = $this->temporaryDirectory . DIRECTORY_SEPARATOR . uniqid();

        if (file_exists($temp)) {
            $this->rmdirRecursive($temp);
        }
        mkdir($temp);

        $zip = new ZipArchive;
        $zip->open($fileHandle);
        $zip->extractTo($temp);
        $zip->close();

        //Check if file exists
        $document = $temp . DIRECTORY_SEPARATOR . 'word' . DIRECTORY_SEPARATOR . 'document.xml';

        if (!file_exists($document)) {
            throw new DocxFileException('Document.xml not found');
        }

        $documentXmlContents = file_get_contents($document);
        $dom = new DOMDocument();
        $loadXMLResult = $dom->loadXML($documentXmlContents, LIBXML_NOERROR | LIBXML_NOWARNING);

        if (!$loadXMLResult || !($dom instanceof DOMDocument)) {
            throw new DocxParsingException("Could not parse XML document");
        }

        return [
            "dom" => $dom,
            "document" => $document,
            "archive" => $temp
        ];
    }

    /**
     * @param $dom
     * @param $archiveLocation
     * @param $saveLocation
     * @throws DocxFileException
     * @internal param $documentXMLLocation
     */
    protected function saveDocument($dom, $archiveLocation, $saveLocation)
    {
        $documentXMLLocation = $archiveLocation . DIRECTORY_SEPARATOR . 'word' . DIRECTORY_SEPARATOR . 'document.xml';
        $newDocumentXMLContents = $dom->saveXml();
        file_put_contents($documentXMLLocation, $newDocumentXMLContents);

        //Create a docx file again
        $zip = new ZipArchive;

        if (!$zip->open($saveLocation, ZipArchive::OVERWRITE)) {
            throw new DocxFileException('Cannot open zip: ' . $saveLocation);
        }

        // Create recursive directory iterator
        $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($archiveLocation), RecursiveIteratorIterator::LEAVES_ONLY);

        foreach($files as $name => $file) {

            $filePath = $file->getRealPath();

            if (in_array($file->getFilename(), array('.', '..'))) {
                continue;
            }

            if (!file_exists($filePath)) {
                throw new DocxFileException('File does not exists: ' . $file->getPathname() . PHP_EOL);
            } else {
                if (!is_readable($filePath)) {
                    throw new DocxFileException('File is not readable: ' . $file->getPathname());
                } else {
                    if (!$zip->addFile($filePath, substr($file->getPathname(), strlen($archiveLocation) + 1))) {
                        throw new DocxFileException('Error adding file: ' . $file->getPathname());
                    }
                }
            }
        }
        if (!$zip->close()) {
            throw new DocxFileException('Could not create zip file');
        }
    }

    /**
     * Helper to remove tmp dir
     *
     * @param $dir
     * @return bool
     */
    protected function rmdirRecursive($dir)
    {
        $files = array_diff(scandir($dir), array('.', '..'));
        foreach($files as $file) {
            (is_dir("$dir/$file")) ? rmdirRecursive("$dir/$file") : unlink("$dir/$file");
        }
        return rmdir($dir);
    }

}