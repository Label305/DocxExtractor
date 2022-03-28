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
    public function getTemporaryDirectory(): string
    {
        return $this->temporaryDirectory;
    }

    /**
     * @param string $temporaryDirectory
     * @return $this
     */
    public function setTemporaryDirectory(string $temporaryDirectory)
    {
        $this->temporaryDirectory = $temporaryDirectory;
        return $this;
    }

    /**
     * Extract file
     * @param string $filePath
     * @throws DocxFileException
     * @throws DocxParsingException
     * @returns array With "document" key, "dom" and "archive" key both are paths. "document" points to the document.xml
     * and "archive" points to the root of the archive. "dom" is the DOMDocument object for the document.xml.
     */
    protected function prepareDocumentForReading(string $filePath): array
    {
        //Make sure we have a complete and correct path
        $filePath = realpath($filePath) ?: $filePath;

        $temp = $this->temporaryDirectory . DIRECTORY_SEPARATOR . uniqid();

        if (file_exists($temp)) {
            $this->rmdirRecursive($temp);
        }
        mkdir($temp);

        $zip = new ZipArchive;
        $opened = $zip->open($filePath);
        if ($opened !== TRUE) {
            throw new DocxFileException( 'Could not open zip archive ' . $filePath . '[' . $opened . ']' );
        }
        $zip->extractTo($temp);
        $zip->close();

        $documentXMLLocation = $this->findDocumentXml($temp);

        $documentXmlContents = file_get_contents($documentXMLLocation);
        $dom = new DOMDocument();
        $loadXMLResult = $dom->loadXML($documentXmlContents, LIBXML_NOERROR | LIBXML_NOWARNING);

        if (!$loadXMLResult || !($dom instanceof DOMDocument)) {
            throw new DocxParsingException( 'Could not parse XML document' );
        }

        return [
            "dom" => $dom,
            "document" => $documentXMLLocation,
            "archive" => $temp
        ];
    }

    /**
     * @param DOMDocument $dom
     * @param string $archiveLocation
     * @param string $saveLocation
     * @throws DocxFileException
     */
    protected function saveDocument(DOMDocument $dom, string $archiveLocation, string $saveLocation)
    {
        if(!file_exists($archiveLocation)) {
            throw new DocxFileException( 'Archive should exist: '. $archiveLocation );
        }

        $documentXMLLocation = $this->findDocumentXml($archiveLocation);

        $newDocumentXMLContents = $dom->saveXml();
        file_put_contents($documentXMLLocation, $newDocumentXMLContents);

        //Create a docx file again
        $zip = new ZipArchive;

        $opened = $zip->open($saveLocation, ZIPARCHIVE::CREATE | ZipArchive::OVERWRITE);
        if ($opened !== true) {
            throw new DocxFileException( 'Cannot open zip: ' . $saveLocation . ' [' . $opened . ']' );
        }

        // Create recursive directory iterator
        $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($archiveLocation), RecursiveIteratorIterator::LEAVES_ONLY);

        foreach($files as $name => $file) {

            $filePath = $file->getRealPath();

            if (in_array($file->getFilename(), array('.', '..'))) {
                continue;
            }

            if (!file_exists($filePath)) {
                throw new DocxFileException( 'File does not exists: ' . $file->getPathname() );
            } else {
                if (!is_readable($filePath)) {
                    throw new DocxFileException( 'File is not readable: ' . $file->getPathname() );
                } else {
                    if (!$zip->addFile($filePath, substr($file->getPathname(), strlen($archiveLocation) + 1))) {
                        throw new DocxFileException( 'Error adding file: ' . $file->getPathname() );
                    }
                }
            }
        }
        if (!$zip->close()) {
            throw new DocxFileException( 'Could not create zip file' );
        }
    }

    /**
     * Helper to remove tmp dir
     *
     * @param string $dir
     * @return bool
     */
    protected function rmdirRecursive(string $dir): bool
    {
        $files = array_diff(scandir($dir), array('.', '..'));
        foreach($files as $file) {
            (is_dir("$dir/$file")) ? $this->rmdirRecursive("$dir/$file") : unlink("$dir/$file");
        }
        return rmdir($dir);
    }

    private function findDocumentXml(string $archiveLocation): ?string
    {
        // Sometimes Word creates another document.xml.
        $possibleFileNames = ['document.xml'];
        for ($i = 0; $i <= 10; $i++) {
            $possibleFileNames[] = sprintf('document%s.xml', $i);
        }

        $document = null;
        foreach ($possibleFileNames as $possibleFileName) {
            //Check if file exists
            $documentpath = $archiveLocation . DIRECTORY_SEPARATOR . 'word' . DIRECTORY_SEPARATOR . $possibleFileName;
            if (file_exists($documentpath)) {
                $document = $documentpath;
            }
        }

        if (!file_exists($document)) {
            throw new DocxFileException( 'document.xml not found' );
        }

        return $document;
    }

}
