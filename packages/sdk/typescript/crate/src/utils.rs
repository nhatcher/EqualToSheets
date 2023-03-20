use equalto_calc::expressions::{
    lexer::util::{get_tokens, MarkedToken},
    token::{OpCompare, OpProduct, OpSum, TokenType},
};

use serde_json::json;
use wasm_bindgen::{prelude::wasm_bindgen, JsError};

trait JsonSerialize {
    fn to_json_string(&self) -> String;
}

impl JsonSerialize for OpSum {
    fn to_json_string(&self) -> String {
        match self {
            OpSum::Add => "Add".to_string(),
            OpSum::Minus => "Minus".to_string(),
        }
    }
}

impl JsonSerialize for OpProduct {
    fn to_json_string(&self) -> String {
        match self {
            OpProduct::Times => "Times".to_string(),
            OpProduct::Divide => "Divide".to_string(),
        }
    }
}

impl JsonSerialize for OpCompare {
    fn to_json_string(&self) -> String {
        match self {
            OpCompare::LessThan => "LessThan".to_string(),
            OpCompare::GreaterThan => "GreaterThan".to_string(),
            OpCompare::Equal => "Equal".to_string(),
            OpCompare::LessOrEqualThan => "LessOrEqualThan".to_string(),
            OpCompare::GreaterOrEqualThan => "GreaterOrEqualThan".to_string(),
            OpCompare::NonEqual => "NonEqual".to_string(),
        }
    }
}

fn to_json_string(marked_token: &MarkedToken) -> String {
    let token = match &marked_token.token {
        TokenType::Illegal(lexer_error) => json!({
            "type":"ILLEGAL",
            "data": {
                "position":lexer_error.position,
                 "message": lexer_error.message.to_string()
            }
        }),
        TokenType::EOF => json!({
            "type": "EOF"
        }),
        TokenType::Ident(s) => json!({
            "type": "IDENT",
            "data": s.to_string()
        }),
        TokenType::String(s) => json!({
            "type": "STRING",
            "data": s.to_string()
        }),
        TokenType::Number(f) => json!({
            "type": "NUMBER",
            "data": *f
        }),
        TokenType::Boolean(b) => json!({
            "type": "BOOLEAN",
            "data": *b
        }),
        TokenType::Error(e) => json!({
            "type": "ERROR",
            "data": e
        }),
        TokenType::Compare(c) => json!({
            "type": "COMPARE",
            "data": c.to_json_string()
        }),
        TokenType::Addition(op) => json!({
            "type": "SUM",
            "data": op.to_json_string()
        }),
        TokenType::Product(op) => json!({
            "type": "PRODUCT",
            "data": op.to_json_string()
        }),
        TokenType::Power => json!({
            "type": "POWER",
        }),
        TokenType::LeftParenthesis => json!({
            "type": "LPAREN",
        }),
        TokenType::RightParenthesis => json!({
            "type": "RPAREN",
        }),
        TokenType::Colon => json!({
            "type": "COLON",
        }),
        TokenType::Semicolon => json!({
            "type": "SEMICOLON",
        }),
        TokenType::LeftBracket => json!({
            "type": "LBRACKET",
        }),
        TokenType::RightBracket => json!({
            "type": "RBRACKET",
        }),
        TokenType::LeftBrace => json!({
            "type": "LBRACE",
        }),
        TokenType::RightBrace => json!({
            "type": "RBRACE",
        }),
        TokenType::Comma => json!({
            "type": "COMMA",
        }),
        TokenType::Bang => json!({
            "type": "BANG",
        }),
        TokenType::Percent => json!({
            "type": "PERCENT",
        }),
        TokenType::And => json!({
            "type": "AND",
        }),
        TokenType::Reference {
            sheet,
            row,
            column,
            absolute_column,
            absolute_row,
        } => {
            let sheet_value = if let Some(sheet_name) = sheet {
                json!(sheet_name)
            } else {
                json!(null)
            };
            json!({
                "type": "REFERENCE",
                "data": {
                    "sheet": sheet_value,
                    "row": row,
                    "column": column,
                    "absolute_column": absolute_column,
                    "absolute_row": absolute_row
                }
            })
        }
        TokenType::Range { sheet, left, right } => {
            let sheet_value = if let Some(sheet_name) = sheet {
                json!(sheet_name)
            } else {
                json!(null)
            };
            json!({
                "type": "RANGE",
                "data": {
                    "sheet": sheet_value,
                    "left": {
                        "row": left.row,
                        "column": left.column,
                        "absolute_column": left.absolute_column,
                        "absolute_row": left.absolute_row
                    },
                    "right": {
                        "row": right.row,
                        "column": right.column,
                        "absolute_column": right.absolute_column,
                        "absolute_row": right.absolute_row
                    }
                }
            })
        }
        TokenType::StructuredReference { .. } => {
            // FIXME
            json!({
                "type": "STRING",
                "data": "".to_string()
            })
        }
    };
    json!({"token": token, "start":marked_token.start,"end":marked_token.end}).to_string()
}

/// Return a JSON string with a list of all the tokens from a formula
/// This is used by the UI to color them according to a theme.
#[wasm_bindgen(js_name = "getFormulaTokens")]
#[allow(dead_code)] // code is not dead, for some reason wasm_bindgen doesn't mark it as in use
pub fn get_formula_tokens(formula: &str) -> Result<String, JsError> {
    let marked_tokens = get_tokens(formula);
    let mut tokens_json: Vec<String> = Vec::new();
    for marked_token in marked_tokens {
        tokens_json.push(to_json_string(&marked_token));
    }
    Ok(format!("[{}]", tokens_json.join(",")))
}
