#![allow(clippy::unwrap_used)]

use crate::{
    expressions::{
        lexer::{Lexer, LexerMode},
        token::{Error, TokenType},
    },
    language::get_language,
    locale::get_locale,
};

fn new_language_lexer(formula: &str, language: &str) -> Lexer {
    let locale = get_locale("en").unwrap();
    let language = get_language(language).unwrap();
    Lexer::new(formula, LexerMode::A1, locale, language)
}

// Spanish

#[test]
fn test_verdadero_falso() {
    let mut lx = new_language_lexer("IF(A1, VERDADERO, FALSO)", "es");
    assert_eq!(lx.next_token(), TokenType::IDENT("IF".to_string()));
    assert_eq!(lx.next_token(), TokenType::LPAREN);
    assert!(matches!(lx.next_token(), TokenType::REFERENCE { .. }));
    assert_eq!(lx.next_token(), TokenType::COMMA);
    assert_eq!(lx.next_token(), TokenType::BOOLEAN(true));
    assert_eq!(lx.next_token(), TokenType::COMMA);
    assert_eq!(lx.next_token(), TokenType::BOOLEAN(false));
    assert_eq!(lx.next_token(), TokenType::RPAREN);
    assert_eq!(lx.next_token(), TokenType::EOF);
}

#[test]
fn test_spanish_errors_ref() {
    let mut lx = new_language_lexer("#¡REF!", "es");
    assert_eq!(lx.next_token(), TokenType::ERROR(Error::REF));
    assert_eq!(lx.next_token(), TokenType::EOF);
}

// German

#[test]
fn test_wahr_falsch() {
    let mut lx = new_language_lexer("IF(A1, WAHR, FALSCH)", "de");
    assert_eq!(lx.next_token(), TokenType::IDENT("IF".to_string()));
    assert_eq!(lx.next_token(), TokenType::LPAREN);
    assert!(matches!(lx.next_token(), TokenType::REFERENCE { .. }));
    assert_eq!(lx.next_token(), TokenType::COMMA);
    assert_eq!(lx.next_token(), TokenType::BOOLEAN(true));
    assert_eq!(lx.next_token(), TokenType::COMMA);
    assert_eq!(lx.next_token(), TokenType::BOOLEAN(false));
    assert_eq!(lx.next_token(), TokenType::RPAREN);
    assert_eq!(lx.next_token(), TokenType::EOF);
}

#[test]
fn test_german_errors_ref() {
    let mut lx = new_language_lexer("#BEZUG!", "de");
    assert_eq!(lx.next_token(), TokenType::ERROR(Error::REF));
    assert_eq!(lx.next_token(), TokenType::EOF);
}

// French

#[test]
fn test_vrai_faux() {
    let mut lx = new_language_lexer("IF(A1, VRAI, FAUX)", "fr");
    assert_eq!(lx.next_token(), TokenType::IDENT("IF".to_string()));
    assert_eq!(lx.next_token(), TokenType::LPAREN);
    assert!(matches!(lx.next_token(), TokenType::REFERENCE { .. }));
    assert_eq!(lx.next_token(), TokenType::COMMA);
    assert_eq!(lx.next_token(), TokenType::BOOLEAN(true));
    assert_eq!(lx.next_token(), TokenType::COMMA);
    assert_eq!(lx.next_token(), TokenType::BOOLEAN(false));
    assert_eq!(lx.next_token(), TokenType::RPAREN);
    assert_eq!(lx.next_token(), TokenType::EOF);
}

#[test]
fn test_french_errors_ref() {
    let mut lx = new_language_lexer("#REF!", "fr");
    assert_eq!(lx.next_token(), TokenType::ERROR(Error::REF));
    assert_eq!(lx.next_token(), TokenType::EOF);
}

// English with errors

#[test]
fn test_english_with_spanish_words() {
    let mut lx = new_language_lexer("IF(A1, VERDADERO, FALSO)", "en");
    assert_eq!(lx.next_token(), TokenType::IDENT("IF".to_string()));
    assert_eq!(lx.next_token(), TokenType::LPAREN);
    assert!(matches!(lx.next_token(), TokenType::REFERENCE { .. }));
    assert_eq!(lx.next_token(), TokenType::COMMA);
    assert_eq!(lx.next_token(), TokenType::IDENT("VERDADERO".to_string()));
    assert_eq!(lx.next_token(), TokenType::COMMA);
    assert_eq!(lx.next_token(), TokenType::IDENT("FALSO".to_string()));
    assert_eq!(lx.next_token(), TokenType::RPAREN);
    assert_eq!(lx.next_token(), TokenType::EOF);
}
