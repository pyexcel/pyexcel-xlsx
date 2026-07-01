#/bin/bash
pip freeze
coverage run -m --source=pyexcel_xlsx pytest --doctest-modules && coverage report --show-missing
