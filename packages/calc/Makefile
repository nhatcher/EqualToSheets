BUILD = cargo build --release --locked
COVERAGE = cargo tarpaulin -v

modules = equalto_calc equalto_xlsx calc

all: bin web

bin:
	$(BUILD)
	mkdir -p build/bin/
	cp target/release/test build/bin/
	cp -r pycalc/pkg/ build/
	cp target/release/libpycalc.so build/pkg/pycalc/_pycalc.so

server: bin
	rm -rf /server/bin
	mkdir -p /server/bin
	cp -r build/bin /server/bin/calc
	rm -rf /server/pkg
	mkdir -p /server/pkg
	cp -r build/pkg /server/

web:
	cd calc && wasm-pack build --target bundler
	cd calc && wasm-pack build --target nodejs --out-dir pkg_node
	cd calc && for i in pkg/*.wasm; do \
			output_file="pkg/$$(basename "$$i" .wasm).opt.wasm"; \
			/usr/bin/wasm-opt "$$i" -o "$$output_file" -O; \
			mv "$$output_file" "$$i"; \
		done
	cd calc && for i in pkg_node/*.wasm; do \
			output_file="pkg_node/$$(basename "$$i" .wasm).opt.wasm"; \
			/usr/bin/wasm-opt "$$i" -o "$$output_file" -O; \
			mv "$$output_file" "$$i"; \
		done
	mkdir -p build/www/pkg
	mkdir -p build/www/pkg_node/
	cp -r calc/pkg/* build/www/pkg/
	cp -r calc/pkg_node/* build/www/pkg_node/

web_dev:
	cd calc && wasm-pack build --dev --target bundler
	cd calc && wasm-pack build --dev --target nodejs --out-dir pkg_node
	mkdir -p build/www/pkg
	mkdir -p build/www/pkg_node/
	cp -r calc/pkg/* build/www/pkg/
	cp -r calc/pkg_node/* build/www/pkg_node/

clean:
	cargo clean
	rm -r -f build/
	rm -r -f target
	rm -r -f calc/pkg
	rm -r -f calc/target
	rm -r -f pycalc/target
	rm -r -f equalto_calc/target
	rm -r -f equalto_xlsx/target

lint:
	cargo fmt -- --check
	# TODO: disallow expect(), unwrap() and panic() via `-D clippy::expect_used -D clippy::unwrap_used -D clippy::panic
	cargo clippy --all-targets --all-features -- -D warnings -A clippy::upper_case_acronyms

format:
	cargo fmt

cargo-tests:
	cargo test --no-default-features

excel-tests:
	cargo test -p equalto_xlsx

tests: cargo-tests excel-tests lint

docs:
	cargo doc --no-deps

coverage:
	cd equalto_calc && $(COVERAGE)
	cd equalto_xlsx && $(COVERAGE)

fetch:
	cargo fetch

.PHONY: lint format cargo-tests docs coverage all
