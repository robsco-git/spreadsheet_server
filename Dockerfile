FROM debian:sid-slim

RUN apt-get update && \
    apt-get upgrade -y && \
    apt-get install -y --no-install-recommends wget python3 python3-dev python3-pip python3-virtualenv python3-uno libreoffice-calc && \
    rm -rf /var/cache/apt/archives /var/lib/apt/lists/*

COPY * /spreadsheet_server/
WORKDIR /spreadsheet_server

RUN mkdir -p /spreadsheet_server/log
RUN virtualenv --system-site-packages -p python3 /opt/virtualenv

ENV PATH="/opt/virtualenv/bin:$PATH"
RUN pip3 install --no-cache-dir -r requirements.txt

CMD ["python3", "-c", "from server import SpreadsheetServer; spreadsheet_server = SpreadsheetServer(host='0.0.0.0'); spreadsheet_server.run()"]