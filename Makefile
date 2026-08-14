PYTHON ?= python3

.PHONY: fetch build test

fetch:
	$(PYTHON) src/fetch_zl.py

build:
	$(PYTHON) src/build_pilot.py --zl sources/ZL3b-n.txt --image-dir assets/folios --repo .

test:
	$(PYTHON) -m unittest discover -s tests -v

