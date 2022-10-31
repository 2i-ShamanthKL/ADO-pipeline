#Python alpine base image
FROM python:3.10.8


ADD . . 

RUN python3 -m pip install pandas

#On startup run the python script

ENTRYPOINT [ "python3", "comparedata.py"]
