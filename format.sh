isort $(find pyexcel_xlsx -name "*.py"|xargs echo) $(find tests -name "*.py"|xargs echo)
black -l 79 pyexcel_xlsx
black -l 79 tests
