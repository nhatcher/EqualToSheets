/*!
# GRAMAR

<pre class="rust">
opComp   => '=' | '<' | '>' | '<=' } '>=' | '<>'
opFactor => '*' | '/'
unaryOp  => '-' | '+'

expr    => concat (opComp concat)*
concat  => term ('&' term)*
term    => factor (opFactor factor)*
factor  => prod (opProd prod)*
prod    => power ('^' power)*
power   => (unaryOp)* range '%'*
range   => primary (':' primary)?
primary => '(' expr ')'
        => number
        => function '(' f_args ')'
        => name
        => string
        => '{' a_args '}'
        => bool
        => bool()
        => error

f_args  => e (',' e)*
</pre>
*/

use crate::language::get_language;
use crate::locale::get_locale;

use super::lexer;
use super::token;
use super::token::OpUnary;
use super::types::*;

use token::OpCompare;
use token::TokenType::*;

pub mod move_formula;
pub mod stringify;

#[cfg(test)]
mod test;

#[cfg(test)]
mod test_ranges;

#[cfg(test)]
mod test_move_formula;

pub(crate) struct Reference<'a> {
    sheet_name: &'a Option<String>,
    sheet_index: u32,
    absolute_row: bool,
    absolute_column: bool,
    row: i32,
    column: i32,
}

#[derive(PartialEq, Clone)]
pub enum Node {
    BooleanKind(bool),
    NumberKind(f64),
    StringKind(String),
    ReferenceKind {
        sheet_name: Option<String>,
        sheet_index: u32,
        absolute_row: bool,
        absolute_column: bool,
        row: i32,
        column: i32,
    },
    RangeKind {
        sheet_name: Option<String>,
        sheet_index: u32,
        absolute_row1: bool,
        absolute_column1: bool,
        row1: i32,
        column1: i32,
        absolute_row2: bool,
        absolute_column2: bool,
        row2: i32,
        column2: i32,
    },
    WrongReferenceKind {
        sheet_name: Option<String>,
        absolute_row: bool,
        absolute_column: bool,
        row: i32,
        column: i32,
    },
    WrongRangeKind {
        sheet_name: Option<String>,
        absolute_row1: bool,
        absolute_column1: bool,
        row1: i32,
        column1: i32,
        absolute_row2: bool,
        absolute_column2: bool,
        row2: i32,
        column2: i32,
    },
    OpRangeKind {
        left: Box<Node>,
        right: Box<Node>,
    },
    OpConcatenateKind {
        left: Box<Node>,
        right: Box<Node>,
    },
    OpSumKind {
        kind: token::OpSum,
        left: Box<Node>,
        right: Box<Node>,
    },
    OpProductKind {
        kind: token::OpProduct,
        left: Box<Node>,
        right: Box<Node>,
    },
    OpPowerKind {
        left: Box<Node>,
        right: Box<Node>,
    },
    FunctionKind {
        name: String,
        args: Vec<Node>,
    },
    ArrayKind(Vec<Node>),
    VariableKind(String),
    CompareKind {
        kind: OpCompare,
        left: Box<Node>,
        right: Box<Node>,
    },
    UnaryKind {
        kind: OpUnary,
        right: Box<Node>,
    },
    ErrorKind(token::Error),
    ParseErrorKind {
        formula: String,
        message: String,
        position: usize,
    },
    EmptyArgKind,
}

#[derive(Clone)]
pub struct Parser {
    lexer: lexer::Lexer,
    worksheets: Vec<String>,
    context: Option<CellReferenceRC>,
}

impl Parser {
    pub fn new(worksheets: Vec<String>) -> Parser {
        let lexer = lexer::Lexer::new(
            "",
            lexer::LexerMode::A1,
            get_locale("en").expect(""),
            get_language("en").expect(""),
        );
        Parser {
            lexer,
            worksheets,
            context: None,
        }
    }
    pub fn set_lexer_mode(&mut self, mode: lexer::LexerMode) {
        self.lexer.set_lexer_mode(mode)
    }

    pub fn set_worksheets(&mut self, worksheets: Vec<String>) {
        self.worksheets = worksheets;
    }

    pub fn parse(&mut self, formula: &str, context: &Option<CellReferenceRC>) -> Node {
        self.lexer.set_formula(formula);
        self.context = context.clone();
        self.parse_expr()
    }

    fn get_sheet_index_by_name(&self, name: &str) -> Option<u32> {
        let worksheets = &self.worksheets;
        for (i, sheet) in worksheets.iter().enumerate() {
            if sheet == name {
                return Some(i as u32);
            }
        }
        None
    }

    fn parse_expr(&mut self) -> Node {
        let mut t = self.parse_concat();
        if let Node::ParseErrorKind { .. } = t {
            return t;
        }
        let mut next_token = self.lexer.peek_token();
        while let COMPARE(op) = next_token {
            self.lexer.advance_token();
            let p = self.parse_concat();
            if let Node::ParseErrorKind { .. } = p {
                return p;
            }
            t = Node::CompareKind {
                kind: op,
                left: Box::new(t),
                right: Box::new(p),
            };
            next_token = self.lexer.peek_token();
        }
        t
    }

    fn parse_concat(&mut self) -> Node {
        let mut t = self.parse_term();
        if let Node::ParseErrorKind { .. } = t {
            return t;
        }
        let mut next_token = self.lexer.peek_token();
        while next_token == AND {
            self.lexer.advance_token();
            let p = self.parse_term();
            if let Node::ParseErrorKind { .. } = p {
                return p;
            }
            t = Node::OpConcatenateKind {
                left: Box::new(t),
                right: Box::new(p),
            };
            next_token = self.lexer.peek_token();
        }
        t
    }

    fn parse_term(&mut self) -> Node {
        let mut t = self.parse_factor();
        if let Node::ParseErrorKind { .. } = t {
            return t;
        }
        let mut next_token = self.lexer.peek_token();
        while let SUM(op) = next_token {
            self.lexer.advance_token();
            let p = self.parse_factor();
            if let Node::ParseErrorKind { .. } = p {
                return p;
            }
            t = Node::OpSumKind {
                kind: op,
                left: Box::new(t),
                right: Box::new(p),
            };

            next_token = self.lexer.peek_token();
        }
        t
    }

    fn parse_factor(&mut self) -> Node {
        let mut t = self.parse_prod();
        if let Node::ParseErrorKind { .. } = t {
            return t;
        }
        let mut next_token = self.lexer.peek_token();
        while let PRODUCT(op) = next_token {
            self.lexer.advance_token();
            let p = self.parse_prod();
            if let Node::ParseErrorKind { .. } = p {
                return p;
            }
            t = Node::OpProductKind {
                kind: op,
                left: Box::new(t),
                right: Box::new(p),
            };
            next_token = self.lexer.peek_token();
        }
        t
    }

    fn parse_prod(&mut self) -> Node {
        let mut t = self.parse_power();
        if let Node::ParseErrorKind { .. } = t {
            return t;
        }
        let mut next_token = self.lexer.peek_token();
        while next_token == POWER {
            self.lexer.advance_token();
            let p = self.parse_power();
            if let Node::ParseErrorKind { .. } = p {
                return p;
            }
            t = Node::OpPowerKind {
                left: Box::new(t),
                right: Box::new(p),
            };
            next_token = self.lexer.peek_token();
        }
        t
    }

    fn parse_power(&mut self) -> Node {
        let mut next_token = self.lexer.peek_token();
        let mut sign = 1;
        while let SUM(op) = next_token {
            self.lexer.advance_token();
            if op == token::OpSum::Minus {
                sign = -sign;
            }
            next_token = self.lexer.peek_token();
        }

        let mut t = self.parse_range();
        if let Node::ParseErrorKind { .. } = t {
            return t;
        }
        if sign == -1 {
            t = Node::UnaryKind {
                kind: token::OpUnary::Minus,
                right: Box::new(t),
            }
        }
        next_token = self.lexer.peek_token();
        while next_token == PERCENT {
            self.lexer.advance_token();
            t = Node::UnaryKind {
                kind: token::OpUnary::Percentage,
                right: Box::new(t),
            };
            next_token = self.lexer.peek_token();
        }
        t
    }

    fn parse_range(&mut self) -> Node {
        let t = self.parse_primary();
        if let Node::ParseErrorKind { .. } = t {
            return t;
        }
        let next_token = self.lexer.peek_token();
        if next_token == COLON {
            self.lexer.advance_token();
            let p = self.parse_primary();
            if let Node::ParseErrorKind { .. } = p {
                return p;
            }
            return Node::OpRangeKind {
                left: Box::new(t),
                right: Box::new(p),
            };
        }
        t
    }

    fn parse_primary(&mut self) -> Node {
        let next_token = self.lexer.next_token();
        match next_token {
            LPAREN => {
                let t = self.parse_expr();
                if let Node::ParseErrorKind { .. } = t {
                    return t;
                }

                if let Err(err) = self.lexer.expect(RPAREN) {
                    return Node::ParseErrorKind {
                        formula: self.lexer.get_formula(),
                        position: err.position,
                        message: err.message,
                    };
                }
                t
            }
            NUMBER(s) => Node::NumberKind(s),
            STRING(s) => Node::StringKind(s),
            LBRACE => {
                let t = self.parse_expr();
                if let Node::ParseErrorKind { .. } = t {
                    return t;
                }
                let mut next_token = self.lexer.peek_token();
                let mut args: Vec<Node> = vec![t];
                while next_token == SEMICOLON {
                    self.lexer.advance_token();
                    let p = self.parse_expr();
                    if let Node::ParseErrorKind { .. } = p {
                        return p;
                    }
                    next_token = self.lexer.peek_token();
                    args.push(p);
                }
                if let Err(err) = self.lexer.expect(RBRACE) {
                    return Node::ParseErrorKind {
                        formula: self.lexer.get_formula(),
                        position: err.position,
                        message: err.message,
                    };
                }
                Node::ArrayKind(args)
            }
            REFERENCE {
                sheet,
                row,
                column,
                absolute_column,
                absolute_row,
            } => {
                let context = match &self.context {
                    Some(c) => c,
                    None => {
                        return Node::ParseErrorKind {
                            formula: self.lexer.get_formula(),
                            position: self.lexer.get_position() as usize,
                            message: "Expected context for the reference".to_string(),
                        }
                    }
                };
                let sheet_index = match &sheet {
                    Some(name) => self.get_sheet_index_by_name(name),
                    None => self.get_sheet_index_by_name(&context.sheet),
                };
                let a1_mode = self.lexer.is_a1_mode();
                let row = if absolute_row || !a1_mode {
                    row
                } else {
                    row - context.row
                };
                let column = if absolute_column || !a1_mode {
                    column
                } else {
                    column - context.column
                };
                match sheet_index {
                    Some(index) => Node::ReferenceKind {
                        sheet_name: sheet,
                        sheet_index: index,
                        row,
                        column,
                        absolute_row,
                        absolute_column,
                    },
                    None => Node::WrongReferenceKind {
                        sheet_name: sheet,
                        row,
                        column,
                        absolute_row,
                        absolute_column,
                    },
                }
            }
            RANGE { sheet, left, right } => {
                let context = match &self.context {
                    Some(c) => c,
                    None => {
                        return Node::ParseErrorKind {
                            formula: self.lexer.get_formula(),
                            position: self.lexer.get_position() as usize,
                            message: "Expected context for the reference".to_string(),
                        }
                    }
                };
                let sheet_index = match &sheet {
                    Some(name) => self.get_sheet_index_by_name(name),
                    None => self.get_sheet_index_by_name(&context.sheet),
                };
                let mut row1 = left.row;
                let mut column1 = left.column;
                let mut row2 = right.row;
                let mut column2 = right.column;

                let mut absolute_column1 = left.absolute_column;
                let mut absolute_column2 = right.absolute_column;
                let mut absolute_row1 = left.absolute_row;
                let mut absolute_row2 = right.absolute_row;

                if self.lexer.is_a1_mode() {
                    if !left.absolute_row {
                        row1 -= context.row
                    };
                    if !left.absolute_column {
                        column1 -= context.column
                    };
                    if !right.absolute_row {
                        row2 -= context.row
                    };
                    if !right.absolute_column {
                        column2 -= context.column
                    };
                }
                if row1 > row2 {
                    (row2, row1) = (row1, row2);
                    (absolute_row2, absolute_row1) = (absolute_row1, absolute_row2);
                }
                if column1 > column2 {
                    (column2, column1) = (column1, column2);
                    (absolute_column2, absolute_column1) = (absolute_column1, absolute_column2);
                }
                match sheet_index {
                    Some(index) => Node::RangeKind {
                        sheet_name: sheet,
                        sheet_index: index,
                        row1,
                        column1,
                        row2,
                        column2,
                        absolute_column1,
                        absolute_column2,
                        absolute_row1,
                        absolute_row2,
                    },
                    None => Node::WrongRangeKind {
                        sheet_name: sheet,
                        row1,
                        column1,
                        row2,
                        column2,
                        absolute_column1,
                        absolute_column2,
                        absolute_row1,
                        absolute_row2,
                    },
                }
            }
            IDENT(name) => {
                let next_token = self.lexer.peek_token();
                if next_token == LPAREN {
                    // It's a function call "SUM(.."
                    self.lexer.advance_token();
                    let args = match self.parse_function_args() {
                        Ok(s) => s,
                        Err(e) => return e,
                    };
                    if let Err(err) = self.lexer.expect(RPAREN) {
                        return Node::ParseErrorKind {
                            formula: self.lexer.get_formula(),
                            position: err.position,
                            message: err.message,
                        };
                    }
                    return Node::FunctionKind {
                        name: name.to_ascii_uppercase(),
                        args,
                    };
                }
                Node::VariableKind(name)
            }
            ERROR(kind) => Node::ErrorKind(kind),
            ILLEGAL(error) => Node::ParseErrorKind {
                formula: self.lexer.get_formula(),
                position: error.position,
                message: error.message,
            },
            EOF => Node::ParseErrorKind {
                formula: self.lexer.get_formula(),
                position: 0,
                message: "Unexpected end of input.".to_string(),
            },
            BOOLEAN(value) => Node::BooleanKind(value),
            COMPARE(_) => {
                // A primary Node cannot start with an operator
                Node::ParseErrorKind {
                    formula: self.lexer.get_formula(),
                    position: 0,
                    message: "Unexpected token: 'COMPARE'".to_string(),
                }
            }
            SUM(_) => {
                // A primary Node cannot start with an operator
                Node::ParseErrorKind {
                    formula: self.lexer.get_formula(),
                    position: 0,
                    message: "Unexpected token: 'SUM'".to_string(),
                }
            }
            PRODUCT(_) => {
                // A primary Node cannot start with an operator
                Node::ParseErrorKind {
                    formula: self.lexer.get_formula(),
                    position: 0,
                    message: "Unexpected token: 'PRODUCT'".to_string(),
                }
            }
            POWER => {
                // A primary Node cannot start with an operator
                Node::ParseErrorKind {
                    formula: self.lexer.get_formula(),
                    position: 0,
                    message: "Unexpected token: 'POWER'".to_string(),
                }
            }
            RPAREN | RBRACKET | COLON | SEMICOLON | RBRACE | COMMA | BANG | AND | PERCENT => {
                Node::ParseErrorKind {
                    formula: self.lexer.get_formula(),
                    position: 0,
                    message: format!("Unexpected token: '{}'", next_token),
                }
            }
            LBRACKET => Node::ParseErrorKind {
                formula: self.lexer.get_formula(),
                position: 0,
                message: "Unexpected token: '['".to_string(),
            },
        }
    }

    fn parse_function_args(&mut self) -> Result<Vec<Node>, Node> {
        let mut args: Vec<Node> = Vec::new();
        let mut next_token = self.lexer.peek_token();
        if next_token == RPAREN {
            return Ok(args);
        }
        if self.lexer.peek_token() == COMMA {
            args.push(Node::EmptyArgKind);
        } else {
            let t = self.parse_expr();
            if let Node::ParseErrorKind { .. } = t {
                return Err(t);
            }
            args.push(t);
        }
        next_token = self.lexer.peek_token();
        while next_token == COMMA {
            self.lexer.advance_token();
            if self.lexer.peek_token() == COMMA {
                args.push(Node::EmptyArgKind);
                next_token = COMMA;
            } else {
                let p = self.parse_expr();
                if let Node::ParseErrorKind { .. } = p {
                    return Err(p);
                }
                next_token = self.lexer.peek_token();
                args.push(p);
            }
        }
        Ok(args)
    }
}
