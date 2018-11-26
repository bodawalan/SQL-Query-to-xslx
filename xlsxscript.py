#!/usr/local/bin/python
import os
import sys
import requests
import json
import xlsxwriter
import mysql.connector
from os.path import join, dirname
from dotenv import load_dotenv
dotenv_path = join(dirname(__file__),'.env')
load_dotenv(dotenv_path)

cnx1 = mysql.connector.connect(user=os.environ['MYSQL_USER'],password = os.environ['MYSQL_PASSWORD'],host = os.environ['MYSQL_HOST'],database=os.environ['MYSQL_DATABASE'])

production = cnx1.cursor();

# run the sql query
def query_sql(cnx1,production,query):
    production.execute(query)
    query_sql.results = production.fetchall();
    query_sql.field_name = [field[0] for field in production.description]
    return ;

# create workbook and work sheet
def Creatworkbook(workbook_name):

    Creatworkbook.workbook = xlsxwriter.Workbook(workbook_name + '.xlsx')
    Creatworkbook.worksheet = Creatworkbook.workbook.add_worksheet();
    return ;
#write data in worksheet based on sql
def Writeworksheetdata(cnx1,production):
    row = 1
    cell_format = Creatworkbook.workbook.add_format({'bold':True})

    header_row =0
    header_col = 0
    for field in query_sql.field_name:
        Creatworkbook.worksheet.write(header_row,header_col,field)
        header_col+=1
    for data in query_sql.results:
        for col in range(len(data)):
            Creatworkbook.worksheet.write(row,col,data[col])
        row+= 1
    Creatworkbook.workbook.close()
    return;


# get sql input from user
def Inputs():
    query = str(input("Write query : ===> "))
    workbook_name = str(input(" write Name for your XLSX file: "))
    Creatworkbook(workbook_name);
    query_sql(cnx1,production,query)
    Writeworksheetdata(cnx1,production)
    return;

Inputs()
