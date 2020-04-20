#!/bin/bash
sudo docker build . -t spreadsheet_server
sudo docker stop spreadsheet_server
sudo docker rm spreadsheet_server
sudo docker run -d \
    -v /home/robsco/code/spreadsheet_server/spreadsheets:/spreadsheet_server/spreadsheets \
    -v /home/robsco/code/spreadsheet_server/saved_spreadsheets:/spreadsheet_server/saved_spreadsheets \
    --restart always \
    -v /home/robsco/code/spreadsheet_server/log:/spreadsheet_server/log \
    -p 5555:5555 \
    --name spreadsheet_server \
    spreadsheet_server
