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
class Paragraph extends ArrayObject
{
    /**
     * @var int
     */
    protected $nextTagIdentifier = 0;

    /**
     * Conenience constructor for the user of the API
     * Strings with <br> <b> <i> <u> <mark> <sub> <sup> and  <font> tags are supported.
     * @param $html string
     * @param Paragraph|null $originalParagraph
     * @return Paragraph
     */
    public static function paragraphWithHTML($html, Paragraph $originalParagraph = null)
    {
        $html = "<html>" . strip_tags($html, '<br /><br><b><strong><em><i><u><mark><sub><sup><font>') . "</html>";
        $html = str_replace("<br>", "<br />", $html);
        $html = str_replace("&nbsp;", " ", $html);
        $htmlDom = new DOMDocument;
        @$htmlDom->loadXml(preg_replace('/&(?!#?[a-z0-9]+;)/', '&amp;', html_entity_decode($html)));

        $paragraph = new Paragraph();
        if ($htmlDom->documentElement !== null) {
            $paragraph->fillWithHTMLDom($htmlDom->documentElement, $originalParagraph);
        }

        return $paragraph;
    }

    /**
     * Recursive method to fill paragraph from HTML data
     *
     * @param DOMNode $node
     * @param Paragraph|null $originalParagraph
     * @param int $br
     * @param bool $bold
     * @param bool $italic
     * @param bool $underline
     * @param bool $highlight
     * @param bool $superscript
     * @param bool $subscript
     * @param bool $hasStyle
     */
    public function fillWithHTMLDom(
        DOMNode $node,
        Paragraph $originalParagraph = null,
        $br = 0,
        $bold = false,
        $italic = false,
        $underline = false,
        $highlight = false,
        $superscript = false,
        $subscript = false,
        $hasStyle = false
    ) {
        if ($node instanceof DOMText) {

            $originalStyle = null;
            if ($originalParagraph !== null) {
                $originalStyle = $this->getOriginalStyle($node, $originalParagraph);
            }
            $this[] = new Sentence($node->nodeValue, $bold, $italic, $underline, $br, $highlight, $superscript, $subscript, $originalStyle);
            $this->nextTagIdentifier++;

        } else {
            if ($node->childNodes !== null) {

                if ($node->nodeName == 'b' || $node->nodeName == 'strong') {
                    $bold = true;
                }
                if ($node->nodeName == 'i' || $node->nodeName == 'em') {
                    $italic = true;
                }
                if ($node->nodeName == 'u') {
                    $underline = true;
                }
                if ($node->nodeName == 'mark') {
                    $highlight = true;
                }
                if ($node->nodeName == 'sup') {
                    $superscript = true;
                }
                if ($node->nodeName == 'sub') {
                    $subscript = true;
                }
                if ($node->nodeName == 'font') {
                    $hasStyle = true;
                }

                foreach ($node->childNodes as $child) {

                    if ($child->nodeName == 'br') {
                        $br++;
                    } else {
                        $this->fillWithHTMLDom($child, $originalParagraph, $br, $bold, $italic, $underline, $highlight, $superscript, $subscript, $hasStyle);
                        $br = 0;
                    }
                }
            }
        }
    }

    /**
     * @param DOMText $node
     * @param Paragraph $originalParagraph
     * @return Style|null
     */
    private function getOriginalStyle(DOMText $node, Paragraph $originalParagraph)
    {
        $originalStyle = null;
        if (array_key_exists($this->nextTagIdentifier, $originalParagraph)) {
            // Sometimes we extract a single space, but in the Paragraph the space is at the beginning of the sentence
            $startsWithSpace = strlen($node->nodeValue) > strlen(ltrim($node->nodeValue));
            if ($startsWithSpace && strlen(ltrim($originalParagraph[$this->nextTagIdentifier]->text)) === 0) {
                // When the current paragraph has no length it may be the space at the beginning
                if (array_key_exists($this->nextTagIdentifier + 1, $originalParagraph)) {
                    // Add the next paragraph style
                    $originalStyle = $originalParagraph[$this->nextTagIdentifier + 1]->style;
                    $this->nextTagIdentifier++;
                }
            } else {
                $originalStyle = $originalParagraph[$this->nextTagIdentifier]->style;
            }
        }
        return $originalStyle;
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
        $highlightActive = false;
        $superscriptActive = false;
        $subscriptActive = false;

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

            $openHighlight = false;
            if ($sentence->highlight && !$highlightActive) {
                $highlightActive = true;
                $openHighlight = true;
            }

            $openSuperscript = false;
            if ($sentence->superscript && !$superscriptActive) {
                $superscriptActive = true;
                $openSuperscript = true;
            }

            $openSubscript = false;
            if ($sentence->subscript && !$subscriptActive) {
                $subscriptActive = true;
                $openSubscript = true;
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

            $closeHighlight = false;
            if ($nextSentence === null || (!$nextSentence->highlight && $highlightActive)) {
                $highlightActive = false;
                $closeHighlight = true;
            }

            $closeSuperscript = false;
            if ($nextSentence === null || (!$nextSentence->superscript && $superscriptActive)) {
                $superscriptActive = false;
                $closeSuperscript = true;
            }

            $closeSubscript = false;
            if ($nextSentence === null || (!$nextSentence->subscript && $subscriptActive)) {
                $subscriptActive = false;
                $closeSubscript = true;
            }

            $result .= $sentence->toHTML($openBold, $openItalic, $openUnderline, $openHighlight, $openSuperscript,
                $openSubscript, $closeBold, $closeItalic, $closeUnderline, $closeHighlight, $closeSuperscript,
                $closeSubscript);
        }

        return $result;
    }
}
