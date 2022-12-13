use crate::expressions::{
    lexer::util::get_tokens,
    token::{OpCompare, OpSum, TokenType},
};

fn get_tokens_types(formula: &str) -> Vec<TokenType> {
    let marked_tokens = get_tokens(formula);
    marked_tokens.iter().map(|s| s.token.clone()).collect()
}

#[test]
fn test_get_tokens() {
    let formula = "1+1";
    let t = get_tokens(formula);
    assert_eq!(t.len(), 3);

    let formula = "1 +   AA23  +";
    let t = get_tokens(formula);
    assert_eq!(t.len(), 4);
    let l = t.get(2).expect("expected token");
    assert_eq!(l.start, 3);
    assert_eq!(l.end, 10);
}

#[test]
fn test_simple_tokens() {
    assert_eq!(
        get_tokens_types("()"),
        vec![TokenType::LPAREN, TokenType::RPAREN]
    );
    assert_eq!(
        get_tokens_types("{}"),
        vec![TokenType::LBRACE, TokenType::RBRACE]
    );
    assert_eq!(
        get_tokens_types("[]"),
        vec![TokenType::LBRACKET, TokenType::RBRACKET]
    );
    assert_eq!(get_tokens_types("&"), vec![TokenType::AND]);
    assert_eq!(
        get_tokens_types("<"),
        vec![TokenType::COMPARE(OpCompare::LessThan)]
    );
    assert_eq!(
        get_tokens_types(">"),
        vec![TokenType::COMPARE(OpCompare::GreaterThan)]
    );
    assert_eq!(
        get_tokens_types("<="),
        vec![TokenType::COMPARE(OpCompare::LessOrEqualThan)]
    );
    assert_eq!(
        get_tokens_types(">="),
        vec![TokenType::COMPARE(OpCompare::GreaterOrEqualThan)]
    );
    assert_eq!(
        get_tokens_types("IF"),
        vec![TokenType::IDENT("IF".to_owned())]
    );
    assert_eq!(get_tokens_types("45"), vec![TokenType::NUMBER(45.0)]);
    // The lexer parses this as two tokens
    assert_eq!(
        get_tokens_types("-45"),
        vec![TokenType::SUM(OpSum::Minus), TokenType::NUMBER(45.0)]
    );
    assert_eq!(
        get_tokens_types("23.45e-2"),
        vec![TokenType::NUMBER(23.45e-2)]
    );
    assert_eq!(
        get_tokens_types("4-3"),
        vec![
            TokenType::NUMBER(4.0),
            TokenType::SUM(OpSum::Minus),
            TokenType::NUMBER(3.0)
        ]
    );
    assert_eq!(get_tokens_types("True"), vec![TokenType::BOOLEAN(true)]);
    assert_eq!(get_tokens_types("FALSE"), vec![TokenType::BOOLEAN(false)]);
    assert_eq!(
        get_tokens_types("2,3.5"),
        vec![
            TokenType::NUMBER(2.0),
            TokenType::COMMA,
            TokenType::NUMBER(3.5)
        ]
    );
    assert_eq!(
        get_tokens_types("2.4;3.5"),
        vec![
            TokenType::NUMBER(2.4),
            TokenType::SEMICOLON,
            TokenType::NUMBER(3.5)
        ]
    );
    assert_eq!(
        get_tokens_types("AB34"),
        vec![TokenType::REFERENCE {
            sheet: None,
            row: 34,
            column: 28,
            absolute_column: false,
            absolute_row: false
        }]
    );
    assert_eq!(
        get_tokens_types("$A3"),
        vec![TokenType::REFERENCE {
            sheet: None,
            row: 3,
            column: 1,
            absolute_column: true,
            absolute_row: false
        }]
    );
    assert_eq!(
        get_tokens_types("AB$34"),
        vec![TokenType::REFERENCE {
            sheet: None,
            row: 34,
            column: 28,
            absolute_column: false,
            absolute_row: true
        }]
    );
    assert_eq!(
        get_tokens_types("$AB$34"),
        vec![TokenType::REFERENCE {
            sheet: None,
            row: 34,
            column: 28,
            absolute_column: true,
            absolute_row: true
        }]
    );
    assert_eq!(
        get_tokens_types("'My House'!AB34"),
        vec![TokenType::REFERENCE {
            sheet: Some("My House".to_string()),
            row: 34,
            column: 28,
            absolute_column: false,
            absolute_row: false
        }]
    );
}
