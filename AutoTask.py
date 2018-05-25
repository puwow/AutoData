#-*- coding:gbk -*-
import os
import requests
import json
import web
import sqlite3
import socket
import urllib
import datetime
from apscheduler.schedulers.background import BackgroundScheduler
import logging
from threading import Thread
import subprocess
from commonUtils import commonUtils

dbname = "AutoData.db"
logging.basicConfig(level=logging.INFO)
CONFIG_FILE = "config.ini"
project_path = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
def initDB():
    #��ʼ�����ݿ�
    sql="""
    create table if not exists packrun(
    'id' integer not null primary key autoincrement,
    'pack_name' varchar(32) not null,
    'pack_desc' varchar(255),
    'pack_url' varchar(128),
    'pack_version' varchar(16) not null,
    'pack_env' varchar(16) not null,
    'pack_type' varchar(1) not null default '0',
    'pack_phone_type' varchar(1) not null default '0',
    'pack_create_date' varchar(32),
    'pack_times' integer default 0,
    'pack_state' varchar(1) not null default '0'
);

create unique index if not exists idx_packrun on packrun('pack_name','pack_version','pack_env','pack_type');
    """
    conn = None
    try:
        conn = sqlite3.connect( dbname )
        conn.executescript(sql)
    except Exception as e:
        raise Exception(e.message)
    finally:
        if conn:
            conn.close()
class PackRunTable:
    def __init__( self ):
        initDB()
    def GetAll( self ):
        db  = web.database(dbn="sqlite", db=dbname )
        ret = db.select("packrun")
        result =  ret.list()
        return result

def checkServer( timeout=3 ):
    #�����ڼ��ĳ��URL�Ƿ�ɷ���
    pack_url = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "pack_url" )
    try:
        socket.setdefaulttimeout(timeout)
        ret = urllib.urlopen( pack_url )
        if ret.code != 200:
            raise Exception(u'�޷����ӷ���!')
    except Exception as e:
        raise Exception(e.message)
    finally:
        socket.setdefaulttimeout(None)

def checkSVN( ):
    svn_url = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "svn_url" )
    try:
        p = subprocess.check_output( "svn info %s"%(svn_url), shell=True )
    except Exception as e:
        raise Exception(e.message)

def checkDevice():
    #����ֻ������Ƿ�����
    logging.info(u'���/�����ֻ��豸BEGIN...')
    ret = commonUtils.Executeable("adb")
    adb_exe = None
    if ret is False:
        adb_exe = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "adb_exe" )
    else:
        adb_exe = "adb"
    if adb_exe and adb_exe !='':
        try:
            p = subprocess.check_output( "%s devices"%(adb_exe), shell=True )
            result = p.split('\n')
            if len( result ) > 1:
                deviceInfo = result[1].split('\t')
                kv={"deviceId":deviceInfo[0].replace("\r","").replace("\n","")}
                logging.info(kv)
                return kv
            else:
                raise Exception(u'δ�ҵ��ֻ��豸')
        except subprocess.CalledProcessError as e:
            raise Exception(u'�����ֻ��豸ʧ��')
        finally:
            logging.info(u'���/�����ֻ��豸END...')
    else:
        logging.info(u'��ȷ�Ͻ�adbָ����ӵ�ϵͳ��������PATH��')
        logging.info(u'���/�����ֻ��豸END...')
        raise Exception(u'��ȷ�Ͻ�adbָ����ӵ�ϵͳ��������PATH��')

def checkKatalon():
    #���Katalon
    katalon_exe = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "katalon_exe" )
    if os.path.exists( katalon_exe ):
        return True
    return False

def updatePack(url, env='VIRT'):
    #���°�
    logging.info(u'���²��԰�BEGIN...')
    new_url = url
    if url[-1] != '/':
        new_url=url+'/'
    try:
        response=None
        response = requests.get(new_url+'runwhat.do?env=%s'%(env) )
        if response is None:
            return
        record = json.loads( response.text )
        if record.get("returncode") == "0":
            data = record.get("data")
            if data:
                pack_name = data.get('pack_name')
                pack_type = data.get('pack_type')
                pack_version = data.get('pack_version')
                pack_phone_type = data.get('pack_phone_type')
                pack_env = data.get('pack_env' )
                pack_create_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                pack_desc = data.get('pack_desc')

                db = web.database(dbn="sqlite", db=dbname)
                ret = db.select("packrun", where="pack_name=$pack_name and pack_type=$pack_type and pack_version=$pack_version and pack_env=$pack_env", vars={"pack_name":pack_name, "pack_type":pack_type, "pack_version":pack_version, "pack_env":pack_env})
                if len(ret.list()) == 0:
                    #���ذ�װ��
                    try:
                        result = requests.get(new_url+"downloads.do?filename=%s&pack_type=%s&pack_phone_type=%s&pack_version=%s&pack_env=%s"%( pack_name, pack_type, pack_phone_type, pack_version, pack_env))
                        if not os.path.exists("downloads"):
                            os.mkdir("downloads")
                        pack_url = os.path.join( os.getcwd(), "downloads", pack_name)
                        with open( pack_url, "wb" ) as fp:
                            for buff in result:
                                fp.write(buff)
                        #��������
                        db.insert("packrun", pack_name=pack_name, pack_type=pack_type, pack_phone_type=pack_phone_type, pack_version=pack_version, pack_env=pack_env, pack_create_date=pack_create_date,\
                                pack_desc=pack_desc, pack_url=pack_url)
                        logging.info(u'����²��԰���%s��'%(pack_name))
                    except Exception as e:
                        raise Exception(u'�����ļ�ʧ��[%s]'%(e.message))
                    logging.info(u'�ҵ����°�[%s]'%(data))
        else:
            logging.info(u'δ��⵽�µİ汾��')
    except Exception as e:
        raise Exception(e.message)
    finally:
        logging.info(u'���²��԰�END...')

def GetQueuePack( pack_type='0', pack_env='VIRT' ):
    #�ŶӴ�ִ���Զ������԰�
    try:
        db = web.database(dbn="sqlite", db=dbname)
        ret = db.select("packrun", where="pack_type=$pack_type and pack_env=$pack_env and pack_state=$pack_state", vars={"pack_type":pack_type,"pack_env":pack_env, "pack_state":"0"})
        result = ret.list()
        return result
    except Exception as e:
        raise Exception(e.message)

def hasRunning( pack_type='0', pack_env='VIRT' ):
    #�Ƿ���������
    try:
        db = web.database( dbn="sqlite", db=dbname )
        ret = db.select("packrun", where="pack_env=$pack_env and pack_type=$pack_type and pack_state=$pack_state", vars={"pack_type":pack_type,"pack_env":pack_env, "pack_state":"1"})
        result = ret.list()
        if len(result) == 0:
            return (False,[])
        else:
            return (True,result)
    except Exception as e:
        raise Exception(e.message)
def InitPackStatus():
    db = web.database( dbn="sqlite", db=dbname )
    try:
        ret = db.update("packrun", where="pack_state=$pack_state", pack_state="0", vars={"pack_state":"1"})
    except Exception as e:
        logging.error(e.message)

def GetRunningPack( self, pack_type='0', pack_env='VIRT' ):
    #��ѯ�������еİ�
    try:
        db = web.database(dbn="sqlite", db=dbname)
        ret = db.select("packrun", where="pack_env=$pack_env and pack_type=$pack_type and pack_state=$pack_state", vars={"pack_type":pack_type,"pack_env":pack_env, "pack_state":"1"})
        result = ret.list()
        if len(result) == 0 :
            ret = db.select("packrun", where="pack_env=$pack_env and pack_type=$pack_type and pack_state=$pack_state", vars={"pack_type":pack_type,"pack_env":pack_env, "pack_state":"0"})
            result = ret.list()
            if( len(result) == 0 ):
                return result
            else:
                ret = db.update("packrun", where="pack_name=$pack_name and pack_type=$pack_type and pack_version=$pack_version", pack_state='1', pack_name=result[0]['pack_name'], \
                        pack_type=result[0]['pack_type'], pack_version=result[0]['pack_version'])
                return GetRunningPack( pack_type, pack_env )
        else:
            return result[0]
    except Exception as e:
        raise Exception(e.message)

def checkoutSvn( url ):
    logging.info(u'SVN��������BEGIN...')
    try:
        p = subprocess.check_output( "svn checkout %s %s"%(url, project_path), shell=True)
        for item in p.split('\n'):
            logging.info(str(item))
    except subprocess.CalledProcessError as e:
        raise Exception(e)
    finally:
        logging.info(u'SVN��������END...')


def setDevice( project_path ):
    #������ĿĬ��ִ�е��ֻ��豸
    logging.info(u'���/�����ֻ��豸BEGIN...')
    ret = commonUtils.Executeable("adb")
    adb_exe = None
    if ret is False:
        adb_exe = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "adb_exe" )
    else:
        adb_exe = "adb"
    if adb_exe and adb_exe !='':
        try:
            p = subprocess.check_output( "%s devices"%(adb_exe), shell=True )
            result = p.split('\n')
            if len( result ) > 1:
                deviceInfo = result[1].split('\t')
                kv={"deviceId":deviceInfo[0]}
                logging.info(kv)
                try:
                    (ret, message) = commonUtils.SetDeviceInfo( project_path, kv )
                    if ret == 0:
                        return kv
                    else:
                        return {}
                except Exception as e:
                    raise Exception(u'���������豸ʧ��')
            else:
                raise Exception(u'δ�ҵ��ֻ��豸')
        except subprocess.CalledProcessError as e:
            raise Exception(u'�����ֻ��豸ʧ��')
        finally:
            logging.info(u'���/�����ֻ��豸END...')
    else:
        logging.info(u'��ȷ�Ͻ�adbָ����ӵ�ϵͳ��������PATH��')
        logging.info(u'���/�����ֻ��豸END...')
        raise Exception(u'��ȷ�Ͻ�adbָ����ӵ�ϵͳ��������PATH��')

def GetDevice():
    adb_exe = None
    ret = commonUtils.Executeable("adb")
    if ret is False:
        adb_exe = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "adb_exe" )
    else:
        adb_exe = "adb"
    if adb_exe and adb_exe !='':
        try:
            p = subprocess.check_output( "%s devices"%(adb_exe), shell=True )
            result = p.split('\n')
            if len( result ) > 1:
                deviceInfo = result[1].split('\t')
                kv={"deviceId":deviceInfo[0].replace("\r","").replace("\n","")}
                return kv
            else:
                raise Exception(u'δ�ҵ��ֻ��豸')
        except subprocess.CalledProcessError as e:
            raise Exception(u'�����ֻ��豸ʧ��')

def clearPack():
    #����Ѿ�ִ�й��Զ������ԵĲ��԰�
    try:
        db = web.database(dbn="sqlite", db="packrun")
        db.delete( "packrun", where="pack_state='1'" )
        logging.info(u'������԰��ɹ�!')
    except Exception as e:
        raise Exception(u'������ɰ�����[%s]'%(e.message))

if __name__ == '__main__':
    initDB()
    suitePath="20180115.ts"
    try:
        updatePack( url="http://192.168.1.222:8080/" )
        checkoutSvn("https://192.168.1.29:8443/svn/branch/branches/spareDir/�ֻ���������/devTools/") 
        kv = setDevice(project_path)
        deviceId = kv.get('deviceId')
        autoRun(project_path, suitePath, deviceId)
    except Exception as e:
        print e.message
