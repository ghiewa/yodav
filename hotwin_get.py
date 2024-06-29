import requests
import random
import io
import time
import datetime
import codecs
import csv
import xlrd,xlwt
import os,shutil
import re
import json
import pandas as pd
import string

from requests import exceptions
from requests.auth import HTTPDigestAuth
from collections import OrderedDict
from pypac import PACSession
from utils.mail_report import io_excel
from datetime import datetime,timedelta,date

from xlutils.copy import copy
from xlrd import open_workbook
from django.core.mail import send_mail, EmailMessage
os.environ['DJANGO_SETTINGS_MODULE'] = 'settings'

RECEIVE_DATA_NOTIFICATION = [
                             'susie.sun@ap.averydennison.com',
                             'sally.du@ap.averydennison.com',
                             'kiko.liu@ap.averydennison.com',
                             'emma.zhao1@ap.averydennison.com'
                            ]
FROMMAIL = 'Hotwind<no-reply@ap.averydennison.com>'

json_bakeup='C:\\Project\\hotwind\\json backup\\'
cs_order="S:\\CSE\\CHCS\\@HOTWIND_FTP\\"

url="http://scm.api.ihotwind.cn/api/pushtagorder/getTagOrders"
s_url="http://scm.api.ihotwind.cn/api/pushtagorder/callbackTagOrders"

def get_order():    

#-----------------------数据获取-------------------------------

    nonce=''.join(random.sample(string.ascii_letters+string.digits,18))
    print (nonce)
    headers = {                
                'Content-Type': 'application/json'
               }
    dd=datetime.now()
    dt=(dd.strftime('%Y-%m-%d %H:%M:%S'))

    print (dt)
    data={
            "time":dt,
            "nonce":nonce,
            "number":100
        }

    data=json.dumps(data)

    try:
        session=PACSession()
        
        r=session.post(url,data=data, headers=headers)
        #print (r)
        #print (r.text)

        if u"无订单数据" not in  r.text:           
            dt=datetime.now()       
            dt=(dt.strftime('%y%m%d%H%M%S'))        
            of = open(json_bakeup+dt+'.json', 'wb')        
            of.write(r.text.encode('utf-8'))
            of.close()

            r_code=r.text[10:13]       
            print (r_code)            
            if r_code == "200" :        
                print (dt)
                order_process(dt)
            else:
                print (r.text)
        else:
            print (r.text)

    except Exception as e:
        print (u"异常: " + str(e))
        pass      

#------------------------------json转EXCEL--------------------------------
def order_process(dt):

    exls=[]
    exl_list=[]

    dt=dt

    with io.open(json_bakeup + dt + '.json', encoding='utf-8') as data_file: 
        all_data=json.load(data_file,object_pairs_hook=OrderedDict) 
        
        data=all_data['data']
        #print data

        bar_list=[]
        matrix=[]
        sku=[]
        all_sku=[]
            
        k=0

        o_styleId=''
        o_supplierNo=''
        for flds in data:            

            supplierNo=flds['supplierNo']
            styleId=flds["styleId"]
            barcode=flds["barCodesOrderNo"]
            sname=flds["supplierName"]
            
            #print ('styleId',styleId)
            #print ('supplierNo',supplierNo)
             
            if 0 == k:
                matrix.append(list(flds.keys()))
                k=k+1
                sku.append(list(flds.values()))
                
            if supplierNo==o_supplierNo and styleId==o_styleId:              
                sku.append(list(flds.values()))

            elif o_supplierNo!='' :        
              
                all_sku=matrix+ sku
                #print (all_sku)

                f_cit=cs_order + o_supplierNo + "-" + o_styleId + "-" +o_sname+'.xls'
                print (f_cit)

                exl_list.append(o_supplierNo + "-" + o_styleId + "-" +o_sname)

                ff=open(f_cit,'wb')
                ff.write(io_excel(all_sku,"Data"))
                ff.close
                all_sku=[]
                sku=[]
                
                sku.append(list(flds.values()))
            o_sname=sname
                
            o_supplierNo=supplierNo
            o_styleId=styleId            
                
            bar_list.append(barcode)
        if o_styleId!='':
            print ("55")
            
            all_sku=matrix+ sku
            #print (all_sku)

            f_cit=cs_order + o_supplierNo + "-" + o_styleId + "-" +o_sname+'.xls'
            print (f_cit)

            ff=open(f_cit,'wb')
            ff.write(io_excel(all_sku,"Data"))
            ff.close
            exl_list.append(o_supplierNo + "-" + o_styleId + "-" +o_sname)   
      
        #print (bar_list)

        headers = {                
                'Content-Type': 'application/json'
               }

        dd=datetime.now()
        dt=(dd.strftime('%Y-%m-%d %H:%M:%S'))

        print (dt)
        s_data={
                "time":dt,
                "barCodesOrderNo":bar_list,               
            }
        s_data=json.dumps(s_data)        
        session=PACSession()
        r=session.post(s_url,data=s_data, headers=headers)
        print (r.text)

        if exl_list:
            body = 'Dear ,<br/><br/>以下订单明细,请及时处理！<br/><br/>'+ json.dumps(exl_list,ensure_ascii=False)      
            subject =  u' Hotwind 新订单'
            msg = EmailMessage(subject, body=body, from_email=FROMMAIL, to=RECEIVE_DATA_NOTIFICATION)
            msg.content_subtype = "html"  # Main content is now text/html
            msg.send()

    
def order(): 
    get_order()
   
if __name__ == '__main__':
    
    order()
