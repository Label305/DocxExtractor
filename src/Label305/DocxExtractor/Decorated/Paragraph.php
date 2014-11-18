<?php namespace Label305\DocxExtractor\Decorated;

use ArrayObject;
use DOMDocument;
use DOMNode;
use DOMText;

/**
 * Class Paragraph
 * @package Label305\DocxExtractor\Decorated
 *
 * Represents a list of <w:r> objects in the docx format. Does not contain
 * <w:p> data. That data is preserved in the extracted document.
 */
class Paragraph extends ArrayObject {

    /**
     * Conenience constructor for the user of the API
     * Strings with <br> <b> <i> and <u> tags are supported.
     * @param $html
     */
    public static function paragraphWithHTML($html)
    {
        $html = "<html>" . strip_tags($html, '<br /><br><b><strong><em><i><u>') . "</html>";
        $htmlDom = new DOMDocument;
        $htmlDom->loadXml($html);

        $paragraph = new Paragraph();
        $paragraph->fillWithHTMLDom($htmlDom->documentElement);
    }

    /**
     * Recursive method to fill paragraph from HTML data
     *
     * @param DOMNode $node
     * @param int $br
     * @param bool $bold
     * @param bool $italic
     * @param bool $underline
     */
    public function fillWithHTMLDom(DOMNode $node, $br = 0, $bold = false, $italic = false, $underline = false)
    {
        if ($node instanceof DOMText) {

            $this[] = new Sentence($node->nodeValue, $bold, $italic, $underline, $br);

        } else if ($node->childNodes !== null) {

            if ($node->nodeName == 'b' || $node->nodeName == 'strong') {
                $bold = true;
            }

            if ($node->nodeName == 'i' || $node->nodeName == 'em') {
                $italic = true;
            }

            if ($node->nodeName == 'u') {
                $underline = true;
            }

            foreach ($node->childNodes as $child) {

                if ($child->nodeName == 'br') {
                    $br++;
                } else {
                    $this->fillWithHTMLDom($child, $br, $bold, $italic, $underline);
                    $br = 0;
                }
            }
        }
    }

    /**
     * Give me a paragraph HTML
     *
     * @return string
     */
    public function toHTML()
    {
        $result = '';

        $boldIsActive = false;
        $italicIsActive = false;
        $underlineIsActive = false;

        for ($i = 0; $i < count($this); $i++) {

            $sentence = $this[$i];

            $openBold = false;
            if ($sentence->bold && !$boldIsActive) {
                $boldIsActive = true;
                $openBold = true;
            }

            $openItalic = false;
            if ($sentence->italic && !$italicIsActive) {
                $italicIsActive = true;
                $openItalic = true;
            }

            $openUnderline = false;
            if ($sentence->underline && !$underlineIsActive) {
                $underlineIsActive = true;
                $openUnderline = true;
            }

            $nextSentence = ($i + 1 < count($this)) ? $this[$i + 1] : null;
            $closeBold = false;
            if ($nextSentence === null || (!$nextSentence->bold && $boldIsActive)) {
                $boldIsActive = false;
                $closeBold = true;
            }

            $closeItalic = false;
            if ($nextSentence === null || (!$nextSentence->italic && $italicIsActive)) {
                $italicIsActive = false;
                $closeItalic = true;
            }

            $closeUnderline = false;
            if ($nextSentence === null || (!$nextSentence->underline && $underlineIsActive)) {
                $underlineIsActive = false;
                $closeUnderline = true;
            }

            $result .= $sentence->toHTML($openBold, $openItalic, $openUnderline, $closeBold, $closeItalic, $closeUnderline);
        }

        return $result;
    }


}