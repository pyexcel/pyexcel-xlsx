pip freeze
nosetests --with-cov --cover-package pyexcel_xlsx --cover-package tests --with-doctest --doctest-extension=.rst tests README.rst pyexcel_xlsx  && flake8 . --exclude=.moban.d 
