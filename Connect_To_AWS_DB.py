#!/usr/bin/env python
# coding: utf-8

# In[1]:


import mysql.connector

# STATIC VARIABLES
HOST = 'database-2.cnkrqj2qjgw2.ap-northeast-1.rds.amazonaws.com'
USERNAME = 'admin'
PASSWORD = 'q2ghGPrmv3PZedw'
PORT = 3306

# Establish connection to AWS
def get_connection():
    mydb = mysql.connector.connect(
                host = HOST,
                user = USERNAME ,
                passwd = PASSWORD,
                port = PORT
            )
    if (mydb):
        mycursor = mydb.cursor(buffered=True)
        return (mydb, mycursor)
    else:
        print("Connection to Database not Estabilished")
        return None






