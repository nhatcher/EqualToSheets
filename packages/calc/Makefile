BUILD = cargo build --release --locked

clean:
	cargo clean
	rm -r -f target
	rm -r -f equalto_calc/target
	rm -r -f equalto_xlsx/target

lint:
	cargo fmt -- --check
	# TODO: disallow expect(), unwrap() and panic() via `-D clippy::expect_used -D clippy::unwrap_used -D clippy::panic
	cargo clippy --all-targets --all-features -- -D warnings -A clippy::upper_case_acronyms

format:
	cargo fmt

coverage:
	cargo llvm-cov --no-default-features --html

excel-tests:
	cargo test -p equalto_xlsx

tests: lint coverage

docs:
	cargo doc --no-deps

fetch:
	cargo fetch

.PHONY: clean lint format coverage tests docs fetch all
