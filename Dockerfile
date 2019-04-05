FROM ubuntu:16.04
RUN apt-get update && apt-get upgrade -y && apt-get install git python3 python3-dev python-virtualenv libreoffice-calc python3-uno gcc python3-pip -y
COPY * /spreadsheet_server/
WORKDIR /spreadsheet_server
RUN pip3 install -r requirements.txt
EXPOSE 5555
CMD ["python3", "-c", "from server import SpreadsheetServer; spreadsheet_server = SpreadsheetServer(host='0.0.0.0'); spreadsheet_server.run()"]