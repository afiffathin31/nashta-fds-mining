@echo off
echo Starting NashTa FDS Mining Edition on http://localhost:5175 ...
start http://localhost:5175
python -m http.server 5175
