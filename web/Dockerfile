FROM python:2.7
WORKDIR /usr/src/app
COPY requirements ./
RUN pip install --no-cache-dir -r requirements
#COPY finance_web/ finance_web/
RUN mkdir finance_web
WORKDIR /usr/src/app/finance_web

CMD ["python", "app.py", "--hostname", "db"]
