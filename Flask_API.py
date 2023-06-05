#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from flask import Flask, request, jsonify, send_file
from sqlalchemy import create_engine, Column, Integer, String, DateTime, Sequence, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime
from sqlalchemy.sql.expression import select, func
import Connect_To_AWS_DB
import json
import uuid
import boto3
import io

app = Flask(__name__)
engine = create_engine('mysql+pymysql://admin:q2ghGPrmv3PZedw@database-2.cnkrqj2qjgw2.ap-northeast-1.rds.amazonaws.com:3306/clevir_db')
Base = declarative_base()

# Define S3 bucket name and AWS credentials
s3_bucket = 'clevirexebucket'
s3_access_key = 'AKIAX5UKROQ2H3EFZVU6'
s3_secret_key = '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'

# Set up S3 client
s3 = boto3.client('s3', aws_access_key_id=s3_access_key, aws_secret_access_key=s3_secret_key)

@app.route('/download_main_exe_file')
def download_main_exe_file():
    s3_object = s3.get_object(Bucket=s3_bucket, Key='main.zip')
    file_stream = io.BytesIO(s3_object['Body'].read())
    file_stream.seek(0)
    return send_file(file_stream, mimetype='application/octet-stream', as_attachment=True, download_name='main.zip')

@app.route('/get_erp_data_key', methods=['GET'])
def get_erp_data_key():
    company_id = request.args.get('company_id')
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    sql = 'SELECT erp_key FROM clevir_db.ErpDataKeys WHERE company_id=%s'
    mycursor.execute(sql, (company_id,))
    erp_data_key = mycursor.fetchone()[0]
    return erp_data_key
    
@app.route('/set_erp_data_key', methods=['POST'])
def set_erp_data_key():
    pass

@app.route('/create_erp_data_key', methods=['POST'])
def create_erp_data_key():
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    pass

@app.route('/all_important_data_dic', methods=['GET'])
def get_first_sync_fix_data_dic():
    company_id = request.args.get('company_id')
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    all_important_data_dic = {}
    all_important_data_dic['SortCodesTypes'] = get_sort_code_types_dic(mydb, mycursor)
    all_important_data_dic['Currencies'] = get_all_currencies_id(mydb, mycursor)
    return json.dumps(all_important_data_dic)

@app.route('/open_new_company', methods=['POST'])
def open_new_company():
    payload = request.get_json()
    new_record_id = str(uuid.uuid4())
    now_creation_date = datetime.utcnow()
    legal_id = int(payload['legal_id'])
    name = str(payload['name'])
    erp_system_id = int(payload['erp_system_id'])
    address = str(payload['address'])
    city = str(payload['city'])
    phone_number = str(payload['phone_number'])
    contact_name = str(payload['contact_name'])
    contact_mobile_phone_number = str(payload['contact_mobile_phone_number'])
    email_address = str(payload['email_address'])
    is_active = int(payload['is_active'])
    book_keeping_firm_id = int(payload['book_keeping_firm_id'])
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    sql = 'INSERT INTO clevir_db.Companies Values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
    values = (new_record_id, now_creation_date, legal_id, name, erp_system_id, address, city, phone_number, contact_name, contact_mobile_phone_number, email_address, is_active, book_keeping_firm_id)
    mycursor.execute(sql, values)
    mydb.commit()
    return jsonify({'token': new_record_id})
    
@app.route('/open_new_user', methods=['POST'])
def open_new_user():
    payload = request.get_json()
    new_record_id = str(uuid.uuid4())
    now_creation_date = datetime.utcnow()
    user_name = str(payload['user_name'])
    password = str(payload['password'])
    main_user_type_id = int(payload['main_user_type_id'])
    is_active = int(payload['is_active'])
    name = str(payload['name'])
    phone_number = str(payload['phone_number'])
    email_address = str(payload['email_address'])
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    sql = 'INSERT INTO clevir_db.Users Values (%s,%s,%s,%s,%s,%s,%s,%s,%s)'
    values = (new_record_id, now_creation_date, user_name, password, main_user_type_id, is_active, name, phone_number, email_address)
    mycursor.execute(sql, values)
    mydb.commit()
    return jsonify({'token': new_record_id})

@app.route('/open_new_user_company', methods=['POST'])
def open_new_user_company():
    payload = request.get_json()
    new_record_id = str(uuid.uuid4())
    now_creation_date = datetime.utcnow()
    company_id = str(payload['company_id'])
    user_id = str(payload['user_id'])
    permissions_type_id = int(payload['permissions_type_id'])
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    sql = 'INSERT INTO clevir_db.UsersCompanies Values (%s,%s,%s,%s,%s)'
    values = (new_record_id, now_creation_date, user_id, company_id, permissions_type_id)
    mycursor.execute(sql, values)
    mydb.commit()
    return jsonify({'token': new_record_id}) 

@app.route('/process_data', methods=['POST'])
def process_data():
    payload = request.get_json()
    company_id = request.args.get('company_id')
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    insert_rows = []
    for key in payload:
        ## We only pass 1 key in the payload each time !!
        if key == 'ErpSortCodes':
            for row_data in payload[key]:
                new_record_id = str(uuid.uuid4())
                now_creation_date = datetime.utcnow().isoformat()
                sort_code, sort_name, sort_code_type_id = row_data
                values = (new_record_id, now_creation_date, company_id, sort_code, sort_name, sort_code_type_id)
                insert_rows.append(values)
            sql = 'INSERT INTO clevir_db.ErpSortCodes VALUES (%s, %s, %s, %s, %s, %s)'
            mycursor.executemany(sql, insert_rows)
            mydb.commit()
            return jsonify(insert_rows)
        elif key == 'ErpAccounts':
            for row_data in payload[key]:
                new_record_id = str(uuid.uuid4())
                now_creation_date = datetime.utcnow()
                erp_sort_code_id, account_number, account_name, account_currency_id, account_creation_date = row_data
                values = (new_record_id, now_creation_date, erp_sort_code_id, account_number, account_name, account_currency_id, account_creation_date)
                insert_rows.append(values)
            sql = 'INSERT INTO clevir_db.ErpAccounts VALUES (%s, %s, %s, %s, %s, %s, %s)'
            mycursor.executemany(sql, insert_rows)
            mydb.commit()
            return jsonify(insert_rows)
        elif key == 'Invoices':
            for row_data in payload[key]:
                new_record_id = str(uuid.uuid4())
                now_creation_date = datetime.utcnow()
                erp_account_id, due_date, reference_date, reference, current_balance_account_currency, amount_account_currency, current_balance_ils, amount_ils, is_open, invoice_type_id = row_data
                values = (new_record_id, now_creation_date, erp_account_id, due_date, reference_date, reference, current_balance_account_currency, amount_account_currency, current_balance_ils, amount_ils, is_open, invoice_type_id)
                insert_rows.append(values)
            sql = 'INSERT INTO clevir_db.Invoices VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'
            mycursor.executemany(sql, insert_rows)
            mydb.commit()
            return jsonify(insert_rows)

@app.route('/update_records_of_company', methods=['POST'])        
def update_records_of_company():
    payload = request.get_json()
    company_id = request.args.get('company_id')
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    erp_sort_id_as_dic = get_erp_sort_id_as_dic_inner_use(mydb, mycursor, company_id)
    erp_accounts_id_and_number_dic = get_erp_accounts_id_and_number_dic_inner_use(mydb, mycursor, company_id)
    updated_rows = []
    for key in payload:
        if key == 'Customers' or key == 'Suppliers':
            for updated_row in payload[key]:
                account_number, reference = updated_row['unique_id']
                account_number = (int(account_number))
                erp_account_id = erp_accounts_id_and_number_dic[account_number]
                fields_changed = updated_row['fields_changed']
                temp_sets = ''
                for field_to_change in fields_changed:
                    field_name, old_value, new_value = field_to_change
                    temp_set = f"{field_name}='{new_value}',"
                    temp_sets += temp_set
                temp_sets = temp_sets[:-1]
                #sql = 'UPDATE clevir_db.Invoices SET %s WHERE erp_account_id=%s AND reference=%s'
                sql = f"UPDATE clevir_db.Invoices SET {temp_sets} WHERE erp_account_id='{erp_account_id}' AND reference='{reference}'"
                #values = (temp_sets, erp_account_id, reference)
                #mycursor.execute(sql, values)
                mycursor.execute(sql)
                mydb.commit()
        elif key == 'Erp Accounts':
            for updated_row in payload[key]:
                account_number, sort_code = (int(updated_row['unique_id'][0])), (int(updated_row['unique_id'][1]))
                erp_sort_code_id = erp_sort_id_as_dic[sort_code]
                fields_changed = updated_row['fields_changed']
                temp_sets = ''
                for field_to_change in fields_changed:
                    field_name, old_value, new_value = field_to_change
                    if field_name == 'account_currency':
                        account_currency_id = find_currency_id_in_dic(get_all_currencies_id(mydb, mycursor), new_value)
                        field_name = 'account_currency_id'
                        new_value = account_currency_id
                    temp_set = f"{field_name}='{new_value}',"
                    temp_sets += temp_set
                temp_sets = temp_sets[:-1]
                sql = f"UPDATE clevir_db.ErpAccounts SET {temp_sets} WHERE erp_sort_code_id='{erp_sort_code_id}' AND account_number='{account_number}'"
                mycursor.execute(sql)
                mydb.commit()
                              
    return {'Status': 200} 
    
@app.route('/get_erp_sort_id_as_dic', methods=['GET'])
def get_erp_sort_id_as_dic():
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    company_id = request.args.get('company_id')
    sql = f'SELECT id, sort_code FROM clevir_db.ErpSortCodes WHERE company_id=%s'
    mycursor.execute(sql,(company_id,))
    erp_sort_code_id_list = mycursor.fetchall()
    final_list = {}
    for row in erp_sort_code_id_list:
        final_list[row[1]] = row[0]
    return final_list  

def get_erp_sort_id_as_dic_inner_use(mydb, mycursor, company_id):
    sql = f'SELECT id, sort_code FROM clevir_db.ErpSortCodes WHERE company_id=%s'
    mycursor.execute(sql,(company_id,))
    erp_sort_code_id_list = mycursor.fetchall()
    final_list = {}
    for row in erp_sort_code_id_list:
        final_list[row[1]] = row[0]
    return final_list 

@app.route('/get_erp_accounts_id_and_number_dic', methods=['GET'])
def get_erp_accounts_id_and_number_dic():
    mydb, mycursor = Connect_To_AWS_DB.get_connection()
    company_id = request.args.get('company_id')
    sql = f'SELECT ErpAccounts.id, ErpAccounts.account_number FROM clevir_db.ErpAccounts INNER JOIN clevir_db.ErpSortCodes ON ErpAccounts.erp_sort_code_id = ErpSortCodes.id WHERE ErpSortCodes.company_id = %s'  
    mycursor.execute(sql,(company_id,))
    erp_sort_code_id_list = mycursor.fetchall()
    final_list = {}
    for row in erp_sort_code_id_list:
        final_list[row[1]] = row[0]
    return final_list

def get_erp_accounts_id_and_number_dic_inner_use(mydb, mycursor, company_id):
    sql = f'SELECT ErpAccounts.id, ErpAccounts.account_number FROM clevir_db.ErpAccounts INNER JOIN clevir_db.ErpSortCodes ON ErpAccounts.erp_sort_code_id = ErpSortCodes.id WHERE ErpSortCodes.company_id = %s'  
    mycursor.execute(sql,(company_id,))
    erp_sort_code_id_list = mycursor.fetchall()
    final_list = {}
    for row in erp_sort_code_id_list:
        final_list[row[1]] = row[0]
    return final_list

### -------------------- Internal use functions -------------------- 

def get_sort_code_types_dic(mydb, mycursor):
    ## get the sort_code_type_id from the database table 'SortCodeTypes'
    sql = 'SELECT * FROM clevir_db.SortCodeTypes'
    mycursor.execute(sql)
    table_data = mycursor.fetchall()
    sort_code_types_dic = {}
    for row in table_data:
        type_id = row[0]
        type_names = []
        for item in range(1, len(row)):
            name = row[item]
            type_names.append(name)
        sort_code_types_dic[type_id] = type_names
    return sort_code_types_dic

def get_all_currencies_id(mydb, mycursor):
    sql = 'SELECT * FROM clevir_db.Currencies'
    mycursor.execute(sql)
    table_data = mycursor.fetchall()
    currencies_dic = {}
    for row in table_data:
        currency_id = row[0]
        currency_names = []
        for item in range(1, len(row)):
            currency_names.append(row[item])
        currencies_dic[currency_id] = currency_names
    return currencies_dic

def get_erp_account_ids_dic(mydb, mycursor, company_id):
    sql = f'SELECT ErpAccounts.id, ErpAccounts.account_number FROM clevir_db.ErpSortCodes JOIN clevir_db.ErpAccounts ON ErpSortCodes.id = ErpAccounts.erp_sort_code_id WHERE company_id=%s'
    mycursor.execute(sql, (company_id,))
    erp_accounts_ids_list = mycursor.fetchall()
    erp_accounts_ids = {}
    for row in erp_accounts_ids_list:
        erp_sort_code_id = row[0]
        account_number = row[1]
        erp_accounts_ids[account_number] = erp_sort_code_id
    return erp_accounts_ids

def find_currency_id_in_dic(currencies_dic, search_currency_name):
    for currency_id in currencies_dic:
        for currency_name in currencies_dic[currency_id]:
            if currency_name == search_currency_name:
                return currency_id
    ## If there's no match of the currency names, we return 1 which is the default
    return 1

if __name__ == '__main__':
    app.run(debug=True)

