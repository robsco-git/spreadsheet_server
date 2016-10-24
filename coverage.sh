#!/bin/bash

# Run the coverage tests and genereate the html reports

#coverage run -m tests.test_connection && coverage html
coverage run -m unittest discover && coverage html
