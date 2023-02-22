define find.functions
	@fgrep -h "##" $(MAKEFILE_LIST) | fgrep -v fgrep | sed -e 's/\\$$//' | sed -e 's/##//'
endef

help:
	@echo 'The following commands can be used.'
	@echo ''
	$(call find.functions)


init: ## sets up environment and installs requirements
init:
	pip install -r requirements.txt

install: ## Installs development requirments
install:
	python -m pip install --upgrade pip
	# Used for packaging and publishing
	pip install -r requirements.txt
	pip install pdf2docx --no-deps
