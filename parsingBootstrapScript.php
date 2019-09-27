<?php

require( __DIR__ . "/vendor/autoload.php" );

$contentDirectory = __DIR__ . "/content";
$outputDirectory = __DIR__ . "/output";
$templateDirectory = __DIR__ . "/src/templates";

$articleFactory = new ArticleConstructor();
$basicRender = new BasicRender();
$articles = [];

if (php_sapi_name() != "cli") {
    header( "Content-type: text/plain" );
}

if ( !is_dir( $contentDirectory ) || !is_readable( $contentDirectory ) ) {
    die( "Content directory unreadable $contentDirectory\n" );
}

// NOTE: disregarding original content folder document order
foreach ( ( new DirectoryIterator( $contentDirectory ) ) as $file ) {
    if ( !$file->isDot() && !$file->isDir() ) {
        if ( $file->getExtension() === 'docx' ) {
            $filePath = "$contentDirectory/" . $file->getFilename();
            try {
                $phpWord = \PhpOffice\PhpWord\IOFactory::load( $filePath, 'Word2007' );
                $article = $articleFactory->fromPHPWord( $phpWord );
                // if title extraction failed document is most probably empty
                if ( !empty( $article->title ) ) {
                    $article->link = md5( $article->title ) . ".html";
                    $articles [] = $article;
                } else {
                    echo "Problem with parsing $filePath - failed to acquire meaningful title\n";
                }
            } catch ( Throwable $throwable ) {
                echo "Problem with parsing $filePath - {$throwable->getMessage()}\n";
            }
        }
    }
}

if ( !is_dir( $outputDirectory ) ) {
    $directoryIsMade = mkdir( $outputDirectory );
    if ( $directoryIsMade === false ) {
        die( "Couldn't create output directory\n" );
//        throw new Exception( "couldn't create output directory" );
    }
}

$rendered = $basicRender->render( "$templateDirectory/list.phtml", [ 'articles' => $articles ] );
$contentWritten = file_put_contents( "$outputDirectory/index.html", $rendered );
if ( $contentWritten === false ) {
    die( "Failed to write output list to index\n" );
//    throw new Exception( "couldn't write output list" );
}
echo "Output list written to $outputDirectory/index.html\n";

foreach ( $articles as $article ) {
    $rendered = $basicRender->render( "$templateDirectory/article.phtml", [ 'article' => $article ] );
    $contentWritten = file_put_contents( "$outputDirectory/$article->link", $rendered );
    if ( $contentWritten === false ) {
        die( "failed to write output into file $article->title" );
//        throw new Exception( "couldn't write output $article->title" );
    }
    echo "Output file written to $outputDirectory/$article->link\n";
}
