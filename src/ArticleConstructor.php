<?php

final class ArticleConstructor
{
    CONST TWIP_TO_INCH_DIVISOR = 1440;

    private function guessTitleInElement ( \PhpOffice\PhpWord\Element\AbstractElement $element ) : ?string {
        $title = null;
        foreach ( $element->getElements() as $element ) {
            $type = get_class( $element );
            switch ( $type ) {
                case \PhpOffice\PhpWord\Element\Title::class:
                    /** @var $element \PhpOffice\PhpWord\Element\Title */
                    $title = $element->getText();
                    break 2;
                case \PhpOffice\PhpWord\Element\Text::class:
                    /** @var $element \PhpOffice\PhpWord\Element\Text */
                    $title = $element->getText();
                    break 2;
                case \PhpOffice\PhpWord\Element\TextRun::class:
                    /** @var $element \PhpOffice\PhpWord\Element\TextRun */
                    $title = "";
                    foreach ( $element->getElements() as $subElement ) {
                        if ( $subElement instanceof \PhpOffice\PhpWord\Element\Text ) {
                            $title .= $subElement->getText();
                        }
                    }
                    break 2;
            }
        }
        if ( is_string( $title ) ) {
            return $title;
        }
        // NOTE: in some cases getText on nodes does not return plain text, but rather another nested node.
        if ( $title instanceof \PhpOffice\PhpWord\Element\TextRun ) {
            $text = "";
            foreach ( $title->getElements() as $subElement ) {
                if ( $subElement instanceof \PhpOffice\PhpWord\Element\Text ) {
                    $text .= $subElement->getText();
                }
            }
            return $text;
        }
        return null;
    }

    private function guessTitle ( \PhpOffice\PhpWord\PhpWord $phpDoc ) {
        $title = null;
        $firstTitle = $phpDoc->getTitles()->getItem( 0 );
        if ( !empty( $firstTitle ) ) {
            $title = $this->guessTitleInElement( $firstTitle );
        }
        if ( empty( $title ) ) {
            foreach ( $phpDoc->getSections() as $section ) {
                $title = $this->guessTitleInElement( $section );
                if ( !empty( $title ) ) {
                    return $title;
                }
            }
        }
        return $title;
    }

    function gatherElementStyle ( $element ) {
        $style = "";
        if ( method_exists( $element, 'getFontStyle' ) ) {
            $font = $element->getFontStyle();
            /** @var $font \PhpOffice\PhpWord\Style\Font */
            if ( !empty( $font ) ) {
                $size = $font->getSize();
                if ( !empty( $size ) ) {
                    $style .= "font-size:{$size}pt;"; // NOTE: Assuming typographic points
                }
                $color = $font->getColor();
                if ( !empty( $color ) ) {
                    $style .= "color:#$color;"; // NOTE: Assuming hex
                }
                if ( !empty( $font->isBold() ) ) {
                    $style .= "font-weight:700;"; // NOTE: Assuming bold as 700 weight
                }
                if ( !empty( $font->isItalic() ) ) {
                    $style .= "font-style:italic;";
                }
            }
        }
        return empty($style) ? null : " style=\"$style\" ";
    }

    function twipToInchRound ( $twipValue ) {
        return round( $twipValue / self::TWIP_TO_INCH_DIVISOR, 2 );
    }

    function gatherSectionStyle ( \PhpOffice\PhpWord\Element\Section $section ) {
        $style = "";
        $sectionStyle = $section->getStyle();
        $width = $this->twipToInchRound( $sectionStyle->getPageSizeW() );
        $style .= "width:{$width}in;";
        $height = $this->twipToInchRound( $sectionStyle->getPageSizeH() );
        $style .= "height:{$height}in;";
        $ml = $this->twipToInchRound( $sectionStyle->getMarginLeft() );
        $mt = $this->twipToInchRound( $sectionStyle->getMarginTop() );
        $mr = $this->twipToInchRound( $sectionStyle->getMarginRight() );
        $mb = $this->twipToInchRound( $sectionStyle->getMarginBottom() );
        $style .= "padding: {$ml}in {$mt}in {$mr}in {$mb}in;";
        $style .= "background:white;";
        return empty($style) ? null : " style=\"$style\" ";
    }

    function populateSectionsRecursive ( $article, $elements ) {
        foreach ( $elements as $element ) {
            $type = get_class( $element );
            $style = $this->gatherElementStyle( $element );
            // TODO: add other element's parsers..
            // NOTE: 'instaceof' requires certain order as elements may implement multiple interfaces
            // TODO: refactor into strategies based on get_class?
            if ( $element instanceof \PhpOffice\PhpWord\Element\Title ) {
                $textNode = $element->getText();
                // NOTE: apparently doc heading levels start from 0, so that's at least +1 level
                // NOTE: I also account for the title extracted separately as another +1
                $depth = min( $element->getDepth(), 5 ) + 2;
                if ( is_string( $textNode ) ) {
                    $article->sections [] = "<h$depth $style>$textNode</h$depth>";
                } else {
                    $article->sections [] = "<h$depth $style>";
                    $this->populateSectionsRecursive( $article, $textNode->getElements() );
                    $article->sections [] = "</h$depth>";
                }
            } elseif ( $element instanceof \PhpOffice\PhpWord\Element\Text ) {
                $textNode = $element->getText();
                if ( is_string( $textNode ) ) {
                    $article->sections [] = "<span $style>$textNode</span>";
                } else {
                    // TODO: it is uncertain if \Text->getText() can also return elements rather than plain text
                    $this->populateSectionsRecursive( $article, $textNode->getElements() );
                }
            } elseif ( $element instanceof \PhpOffice\PhpWord\Element\ListItemRun ) {
                // TODO: \ListItemRun should only make up <ul> wrap, should check why it may not contain \ListItem[]
                $article->sections [] = "<ul $style><li $style>";
                $this->populateSectionsRecursive( $article, $element->getElements() );
                $article->sections [] = "</li></ul>";
            } elseif ( $element instanceof \PhpOffice\PhpWord\Element\TextRun ) {
                $article->sections [] = "<p $style>";
                $this->populateSectionsRecursive( $article, $element->getElements() );
                $article->sections [] = "</p>";
            } elseif ( $element instanceof \PhpOffice\PhpWord\Element\TextBreak ) {
                $article->sections [] = "<hr>";
            } else {
                $article->sections [] = "<span style='color:red'>$type not supported!</span><br>";
            }
        }
    }

    function fromPHPWord ( \PhpOffice\PhpWord\PhpWord $phpWord ) {
        $article = new ArticleDTO();
        $article->title = $this->guessTitle( $phpWord );
        $sections = $phpWord->getSections();
        foreach ( $sections as $section ) {
            $style = $this->gatherSectionStyle( $section );
            $article->sections [] = "<section $style>";
            $this->populateSectionsRecursive( $article, $section->getElements() );
            $article->sections [] = "</section>";
        }
        return $article;
    }
}