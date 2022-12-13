# EqualTo Spreadsheet Technology

# Dependencies and installation

You will need to install Rust and wasm-pack

1. Install Rust
2. Install wasm-pack
3. Build the Rust code

## Install Rust and wasm-pack

To install rust and wasm-pack, follow the instructions in the "Install rust and wasm-pack" section of the [Quickstart](https://github.com/EqualTo-Software/EqualTo/wiki/Getting-started) guide.

## Build everything

This will create a `build` folder and

- Build the Rust equalto_calc crate
- Build the Rust equalto_xlsx crate
- Build the wasm calc package ready for use in the browser
- Build the binaries:
  - eval_workbook (see documenation in the carte)
  - test (test a particular Excel workbook against the Rust implementation)
  - import (imports an xlsx file into an EqualTo json formart)
- In `build/wwww` you can run a simple example server

```bash
equalto/rust$ make
```

### test

To test a particular excel file:

```bash
$ test <input-file>
```

## Tests

Top run all the tests, simply:

```bash
equalto_calc$ make tests
```

This runs:

- Unit tests
- Documentation tests
- Integration tests

To see some example of unit tests look for a `test.rs` file in `equalto_calc/src/parser/parser/test.rs`.

The parser and lexer are heavily tested.

At the time of writing there is only one documentation test in `equalto_xlsx/src/utils.rs`
They are equivalent to unit test, but they show up in the documentation.

Integration tests are located in the `tests` folder of every crate.
For instance in `equalto_xlsx/tests` we test that importing `example.xlsx` produces `example.json`.

# Docs

You can generate some minimal documentation by

```bash
$ make docs
```

Please refer to the cargo (book)[https://doc.rust-lang.org/cargo/commands/cargo-doc.html] on how to see this documentation.

# Test coverage

You need to install `cargo-tarpaulin`

```bash
$ cargo install cargo-tarpaulin
```

Then simply

```bash
$ make coverage
```
