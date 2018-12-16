FROM python:3.6
ADD xlsxscript.py
COPY requirements.txt /opt/app/requirements.txt
WORKDIR /opt/app
RUN pip install -r requirements.txt
ENTRYPOINT ["python","/xlsxscript.py"]
