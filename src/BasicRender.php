<?php

final class BasicRender
{
    static function escape ( string $string ) {
        return htmlspecialchars( $string, ENT_QUOTES, 'UTF-8' );
    }

    function render ( $view, $variables ) {
        ob_start();
        extract( $variables );
        require( $view );
        return ob_get_clean();
    }
}