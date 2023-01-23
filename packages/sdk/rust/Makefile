tests: lint coverage

lint:
	cargo fmt -- --check
	cargo clippy --all-targets --all-features -- -D warnings

fmt:
	cargo fmt

coverage:
	cargo llvm-cov --no-default-features --html

clean:
	cargo clean
	rm -r -f target

.PHONY: clean lint format coverage tests
