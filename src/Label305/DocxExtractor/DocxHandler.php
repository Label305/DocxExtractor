<?php namespace Label305\DocxExtractor;


use DOMDocument;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use ZipArchive;

abstract class DocxHandler {

    /**
     * Defaults to sys_get_temp_dir()
     * 
     * @var string the tmp dir location
     */
    protected $temporaryDirectory;

    /**
     * Sets the temporary directory to the system
     */
    function __construct()
    {
        $this->setTemporaryDirectory(sys_get_temp_dir());
    }
    
    /**
     * @return string
     */
    public function getTemporaryDirectory()
    {
        return $this->temporaryDirectory;
    }

    /**
     * @param string $temporaryDirectory
     * @return $this
     */
    public function setTemporaryDirectory($temporaryDirectory)
    {
        $this->temporaryDirectory = $temporaryDirectory;
        return $this;
    }

    /**
     * Extract file
     * @param $filePath
     * @throws DocxFileException
     * @throws DocxParsingException
     * @returns array With "document" key, "dom" and "archive" key both are paths. "document" points to the document.xml
     * and "archive" points to the root of the archive. "dom" is the DOMDocument object for the document.xml.
     */
    protected function prepareDocumentForReading($filePath)
    {
        //Make sure we have a complete and correct path
        $filePath = realpath($filePath);
        
        $temp = $this->temporaryDirectory . DIRECTORY_SEPARATOR . uniqid();

        if (file_exists($temp)) {
            $this->rmdirRecursive($temp);
        }
        mkdir($temp);

        $zip = new ZipArchive;
        $zip->open($filePath);
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
     */
    protected function saveDocument($dom, $archiveLocation, $saveLocation)
    {
        if(!file_exists($archiveLocation)) {
            throw new DocxFileException('Archive should exist: '. $archiveLocation);
        }
        
        $documentXMLLocation = $archiveLocation . DIRECTORY_SEPARATOR . 'word' . DIRECTORY_SEPARATOR . 'document.xml';
        $newDocumentXMLContents = $dom->saveXml();
        file_put_contents($documentXMLLocation, $newDocumentXMLContents);

        //Create a docx file again
        $zip = new ZipArchive;

        $opened = $zip->open($saveLocation, ZIPARCHIVE::CREATE | ZipArchive::OVERWRITE);
        if ($opened !== true) {
            throw new DocxFileException('Cannot open zip: ' . $saveLocation . ' [' . $opened . ']');
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