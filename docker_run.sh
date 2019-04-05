#!/bin/bash
sudo docker build . -t spreadsheet_server
sudo docker stop spreadsheet_server
sudo docker rm spreadsheet_server
sudo docker run -d -v /home/robsco/Work/spreadsheet_server/spreadsheets:/spreadsheet_server/spreadsheets \
     --restart always \
     -v /home/robsco/Work/spreadsheet_server/log:/spreadsheet_server/log \
     -p 5555:5555 \
     --name spreadsheet_server \
     spreadsheet_server
