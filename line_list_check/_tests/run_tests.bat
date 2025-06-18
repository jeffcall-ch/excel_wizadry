@echo off
echo Running tests with coverage report...
cd %~dp0
pytest -xvs --cov=.. --cov-report=term --cov-report=html
echo Test execution completed.
pause
