#-*- coding:gbk -*-

import os
import re
import json
import fnmatch
import xml.dom.minidom
import xlrd
import xlwt
import chardet
import sys
import ConfigParser
from xlutils.copy import copy
import uuid
from xml.dom.minidom import Document
from  xml.etree import ElementTree
import logging
import subprocess
reload(sys)
sys.setdefaultencoding("gbk")

CONFIG_FILE = "config.ini"
class commonUtils:
    '''
    默认查询所有Test Case/simpCase，并返回迭代器
    { "account":[
            {
            "name":"",
            "desc":"",
            "variables":[ { "name":"", "desc":"", "default":"" } ]
            }
        ]
    }
    '''
    
    @classmethod
    def find_test_cases( cls, simple=True ):
        #绝对路径
        PROJECT_PATH = cls.ConfigRead( "config.ini", "MBKAUTOTEST", "project_path" )
        logging.info(PROJECT_PATH)
        if PROJECT_PATH == None:
            return
        find_path = os.path.join( PROJECT_PATH, "Test Cases" )
        logging.info(find_path)
        regx = None
        if simple:
            find_path = os.path.join( PROJECT_PATH, "Test Cases", "simpCase" )
            regx = re.compile(r"simpCase/*.tc$")
        else:
            regx = re.compile(r"*.tc$")
        records = {}

        logging.info(find_path)
        for root, dirs, files in os.walk( find_path ):
            for subdir in dirs:
                records[subdir]=[]
            record={}
            for fitem in fnmatch.filter(files, "*.tc"):
                record = cls.find_test_case_info( os.path.join( root, fitem ) )
                record['path'] = os.path.join( root, fitem )
                if root != find_path:
                    records[os.path.basename(root)].append( record )
        return records

    @classmethod
    def find_test_case_info( cls, caseName ):
        #获取TestCase的信息，以字典的形式返回
        #caseName是绝对路径
        dom = xml.dom.minidom.parse( caseName )
        doc = dom.documentElement
        test_case_info = {}
        test_case_info['name'] = ""
        test_case_info['desc'] = ""
        for node in doc.childNodes:
            if node.nodeName == "name":
                test_case_info['name'] = node.childNodes[0].nodeValue
            if node.nodeName == "description":
                if node.hasChildNodes():
                    test_case_info['desc'] = node.childNodes[0].nodeValue
                else:
                    test_case_info['desc'] = u"未标记"
        return test_case_info

    @classmethod
    def parse_test_case_2( cls, test_case ):
        #用ElementTree解析TestCase
        #test_case - TestCase绝对路径
        #数据结构{'description':'','name':'','testCaseGuid':'', 'variables':[(name, description, defaultValue, id, masked)]}
        data={}
        testCaseId = None
        g =  re.match( r'.*(Test Cases.*).tc$', test_case )
        if g:
           testCaseId = g.groups()[0]
        try:
            et = ElementTree.parse( test_case )
            root = et.getroot()
            if root is not None:
                desc=root.find("description").text
                name=root.find("name").text
                guid =root.find("testCaseGuid").text
                variables = root.findall("variable")
                data['name']=name
                data['description']=desc
                data['variables']=[]
                data['testCaseGuid']=guid
                data['testCaseId']=testCaseId.replace("\\","/")
                for var in variables:
                    var_name = var.find('name').text
                    var_masked = var.find("masked").text
                    var_description = var.find("description").text
                    var_defaultValue = var.find("defaultValue").text
                    var_id = var.find("id").text
                    data['variables'].append((var_name,var_description, var_defaultValue,var_id,var_masked))
        except Exception as e:
            logging.error(u'解析案例文件[%s]失败[%s]!'%(test_case,e.message) )
            return(-1, u'解析TestCase失败[%s][%s]'%(test_case, e.message) )
        return (0, data)

    @classmethod
    def update_test_case( cls, caseName, row=1 ):
        #更新案例，绑定默认数据
        #caseName - 案例的绝对路径
        #row - 案例默认值是第几行

        try:
            tree = ElementTree.parse( caseName )
            root = tree.getroot()
            #数据名是案例名+Data，如Login.tc的数据文件名是LoginData
            dataname = os.path.basename(caseName).replace(".tc", "Data")
            col = 1
            for element in root.findall('variable'):
                node = element.find("defaultValue")
                if node is not None:
                    node.text="findTestData('%s').getValue(%d, %d)"%(dataname, col, row )
                col=col+1
            doc = xml.dom.minidom.parseString( xml.etree.ElementTree.tostring( root ).replace("\n","") )
            with open( caseName, "w" ) as fp:
                fp.write( doc.toprettyxml( indent="\t", newl='\n', encoding="utf-8") )
        except Exception as e:
            logging.error(u'解析案例文件[%s]失败!'%(e.message))
            return (-1, u'解析案例文件失败[%s]'%(e.message))
        return (0, u'更新案例[%s]成功'%(caseName))

    @classmethod
    def set_style( cls, name, height, bold=False ):
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = name
        font.bold = bold
        font.color_index = 4
        font.height = height
        style.font = font
        return style

    @classmethod
    def init_xls( cls, filename, sheet=True ):
        if sheet:
            wb = xlwt.Workbook()
        for page in cls.find_test_cases():
            page_data = cls.parse_test_case( page )
            if page_data:
                sheet = wb.add_sheet( page_data.get("desc")+"("+page_data.get('name')+")", cell_overwrite_ok=True )
                row0 = [u"编号"]+[ x.get('desc')+"("+x.get('name')+")" for x in page_data.get('variables') ]
                for i in range(len(row0)):
                    sheet.write( 0, i, row0[i], cls.set_style('Times New Roman', 220, True) )
        wb.save(filename)

    @classmethod
    def get_excel_file( cls, sheetname ):
        #去读哪个Excel文件
        data_path = cls.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "data_path" )
        data_name = cls.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "data_name" )
        single_excel = cls.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "single_excel" )

        filename = os.path.join( data_path, data_name )
        if eval(single_excel) is False:
            filename = os.path.join( data_path, sheetname+".xls" )
        return filename

    @classmethod
    def get_deal_sheets( cls, sheetname ):
        #返回所有的sheetname列表，为什么有这一步，因为Excel可能是单文件，也可能是多个文件
        filename = cls.get_excel_file( sheetname )
        result = []
        try:
            if os.path.exists( filename ):
                rb = xlrd.open_workbook( filename )
                for sheet in rb.sheet_names():
                    result.append(sheet)
        except Exception as e:
            logging.error( u'解析Excel文件[%s]失败!'%(e.message) )
        finally:
            return result

    @classmethod
    def read_excel( cls, sheetname ):
        #根据标签页读取相应Excel表中的数据
        data = []
        #获取Excel文件
        sheets = cls.get_deal_sheets( sheetname )
        filename = cls.get_excel_file( sheetname )
        try:
            if os.path.exists( filename ):
                rb = xlrd.open_workbook(filename)
                if sheetname in sheets:
                    sh = rb.sheet_by_index(sheets.index(sheetname))
                    if sh:
                        for rownum in range( sh.nrows ):
                            data.append(sh.row_values( rownum ))
        except Exception as e:
            logging.error(u'读取Excel文件[%s]失败!'%(e.message))
        finally:
            return data

    @classmethod
    def update_excel( cls, sheetname, data ):
        #更新Excel表数据
        filename = cls.get_excel_file( sheetname )
        sheets = cls.get_deal_sheets( sheetname )
        sh = None
        wb = None
        rb = None
        try:
            if os.path.exists( filename ):
                #如果Excel文件已存在
                rb = xlrd.open_workbook( filename )
                wb = copy(rb)
                if sheetname in sheets:
                    #如果Sheet也已经存在
                    sh = wb.get_sheet(sheets.index(sheetname))
                else:
                    sh = wb.add_sheet( data.get('name'), cell_overwrite_ok=True )
            else:
                wb = xlwt.Workbook()
                sh = wb.add_sheet( data.get('name'), cell_overwrite_ok=True )
            if data:
                if len(data.get('data'))>0:
                    row=0
                    for rowData in data.get('data'):
                        col=0
                        if len(rowData) == 0:
                            sh.write( 0, 0, '' )
                        else:
                            for colData in rowData:
                                sh.write( row, col, colData )
                                col=col+1
                            row=row+1
                else:
                    sh.write( 0, 0, '')
            else:
                sh.write( 0, 0, '')
        except Exception as e:
            logging.error(u'更新Excel数据[%s]失败!'%(e.message) )
            return (-1,u'更新Excel数据失败[%s]'%(e.message))
        finally:
            wb.save(filename)
        return (0, u'更新Excel[%s]数据成功'%(filename) )

    @classmethod
    def create_data_xml_2( cls, sheetname ):
        #根据Excel的Sheet页创建Katalon使用的数据文件
        filename = cls.get_excel_file( sheetname )
        dataname = os.path.join( cls.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "data_path" ), sheetname+"Data.dat" )
        try:
            root = ElementTree.Element("DataFileEntity")
            son_desc = ElementTree.SubElement( root, "description" )
            son_name = ElementTree.SubElement( root, "name" )
            son_name.text=sheetname+"Data"
            son_tag = ElementTree.SubElement( root, "tag" )
            son_head = ElementTree.SubElement( root, "containsHeaders" )
            son_head.text="true"
            son_sep = ElementTree.SubElement( root, "csvSeperator" )
            son_data_file = ElementTree.SubElement( root, "dataFile" )
            son_data_file.text=str(uuid.uuid1())
            son_dsu = ElementTree.SubElement( root, "dataSourceUrl" )
            son_dsu.text= filename
            son_driver = ElementTree.SubElement( root, "driver" )
            son_driver.text="ExcelFile"
            son_interP = ElementTree.SubElement( root, "isInternalPath" )
            son_interP.text="false"
            son_query = ElementTree.SubElement( root, "query" )
            son_secUA = ElementTree.SubElement( root, "secureUserAccount" )
            son_secUA.text="false"
            son_sheetName = ElementTree.SubElement( root, "sheetName" )
            son_sheetName.text=sheetname
            son_global = ElementTree.SubElement( root, "usingGlobalDBSetting" )
            son_global.text="false"

            doc = xml.dom.minidom.parseString( xml.etree.ElementTree.tostring( root ) )
            with open( dataname, "w") as fp:
                fp.write( doc.toprettyxml( indent="\t", newl="\n", encoding="utf-8") )
        except Exception as e:
            logging( u'创建映射文件[%s]失败!'%(e.message) )
            return (-1, u'创建映射文件失败[%s]'%(e.message))
        return (0, u'创建映射文件[%s]成功'%(dataname) )
    @classmethod
    def create_data_xml( cls, sheetname, data ):
        #数据文件Excel
        filename = cls.get_excel_file( sheetname )
        #待生成数据文件
        dataname = os.path.join( cls.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "data_path" ), sheetname+"Data.dat" )
        try:
            xmlDoc = Document()
            document = xmlDoc.createElement("DataFileEntity")
            xmlDoc.appendChild( document )

            node = xmlDoc.createElement("description")
            document.appendChild( node )

            node = xmlDoc.createElement( "name" )
            nodeText = xmlDoc.createTextNode(sheetname+"Data")
            node.appendChild(nodeText)
            document.appendChild( node )

            node = xmlDoc.createElement("tag")
            document.appendChild( node )

            node = xmlDoc.createElement("containsHeaders" )
            nodeText = xmlDoc.createTextNode("true")
            node.appendChild(nodeText)
            document.appendChild(node)

            node = xmlDoc.createElement( "csvSeperator" )
            document.appendChild(node)

            node = xmlDoc.createElement("dataFile" )
            nodeText=xmlDoc.createTextNode(str(uuid.uuid1()))
            node.appendChild( nodeText )
            document.appendChild( node )

            node = xmlDoc.createElement("dataSourceUrl")
            nodeText = xmlDoc.createTextNode( filename )
            node.appendChild( nodeText )
            document.appendChild( node )

            node = xmlDoc.createElement( "driver" )
            nodeText = xmlDoc.createTextNode("ExcelFile")
            node.appendChild( nodeText )
            document.appendChild( node )

            node = xmlDoc.createElement( "isInternalPath" )
            nodeText = xmlDoc.createTextNode( "false" )
            node.appendChild( nodeText )
            document.appendChild( node )

            node = xmlDoc.createElement( "query" )
            document.appendChild( node )

            node = xmlDoc.createElement("secureUserAccount")
            nodeText = xmlDoc.createTextNode( "false" )
            node.appendChild( nodeText )
            document.appendChild( node )

            node = xmlDoc.createElement("sheetName" )
            nodeText = xmlDoc.createTextNode( sheetname )
            node.appendChild( nodeText )
            document.appendChild( node )

            node = xmlDoc.createElement("usingGlobalDBSetting")
            nodeText = xmlDoc.createTextNode( "false" )
            node.appendChild( nodeText )
            document.appendChild( node )

            with open( dataname, "w" ) as fp:
                xmlDoc.writexml( fp, indent="\t", newl="\n", encoding="utf-8" )

        except Exception as e:
            logging.error(u"写文件[%s]失败!"%(dataname) )

    @classmethod
    def parse_test_data_2( cls, test_case ):
        #解析TestData
        dataPath = cls.ConfigRead(CONFIG_FILE, "MBKAUTOTEST", "DATA_PATH" )
        if dataPath:
            dataFile = os.path.join( dataPath, os.path.basename(test_case).replace(".tc", "Data.dat" ) )
            if os.path.exists( dataFile ):
                data = {}
                try:
                    et = ElementTree.parse( dataFile )
                    root = et.getroot()
                    node = root.find("name")
                    if node is not None:
                        data['name'] = node.text
                    node = root.find("dataFile")
                    if node is not None:
                        data['dataFile']=node.text
                    node = root.find("sheetName")
                    if node is not None:
                        data['sheetName']=node.text
                    node = root.find("containsHeaders")
                    if node is not None:
                        data['containHeaders'] = node.text
                    node = root.find("dataSourceUrl")
                    if node is not None:
                        data['dataSourceUrl'] = node.text
                except Exception as e:
                    logging.error(u'解析TestData失败[%s]!'%(e.message) )
                    return (-3, u'解析TestData失败[%s]'%(e.message) )
                return (0, data )
            else:
                logging.error(u'数据文件不存在[%s]!'%(datafile) )
                return (-2, u'数据文件不存在[%s]'%(dataFile) )
        else:
            logging.error(u'未设置数据存储目录[%s]'%(dataPath) )
            return (-1, u'未设置数据存储目录[%s]'%(dataPath))

    @classmethod
    def get_case_data_vars( cls, suiteInfo, testCases ):
        #根据案例获取案例关联数据、关联变量等
        #testCases案例列表，需要全路径
        data ={}
        if suiteInfo:
            data={
                "suiteName":suiteInfo.get('suiteName'),
                "name":suiteInfo.get("name"),
                "description":suiteInfo.get("description"), 
                "isRerun":"false", 
                "lastRun":suiteInfo.get("lastRun"), 
                "numberOfRerun":suiteInfo.get("numberOfRerun"),
                "pageLoadTimeout":suiteInfo.get("pageLoadTimeout"), 
                "pageLoadTimeoutDefault":suiteInfo.get("pageLoadTimeoutDefault"), 
                "rerunFailedTestCasesOnly":suiteInfo.get("returnFailedTestCasesOnly") ,
                "testCases":[],
            }
        for testCase in testCases:
            #数据结构{'description':'','name':'','testCaseGuid':'', 'variables':[(name, description, defaultValue, id, masked)]}
            ret = -1
            ret2 = -1
            try:
                ret, caseInfo = cls.parse_test_case_2( testCase )
                ret2,testDataInfo = cls.parse_test_data_2(testCase)
            except Exception as e:
                logging.error(u"获取TestCase信息失败[%s]!"%(e.message))
                return
            if ret == 0 and ret2 == 0:
                testCaseInfo = {
                        #"guid":caseInfo.get("testCaseGuid"),
                        "guid":str(uuid.uuid1()),
                        "isReuseDriver":"false",
                        "isRun":"true",
                        "testCaseId":caseInfo.get('testCaseId'),
                        "testDatas":[],
                        "variables":[]
                }

                for variable in caseInfo.get('variables'):
                    desc=""
                    if variable[1]:
                        desc=variable[1]
                    testCaseInfo.get("variables").append({
                                "testDataLinkId":testDataInfo.get("dataFile"),
                                "type":"DATA_COLUMN",
                                "value":variable[0],
                                "desc":desc,
                                "variableId":variable[3],
                                })
                
                testCaseInfo.get("testDatas").append({
                            'combinationType':'ONE',
                            #'id':testDataInfo.get("dataFile"),
                            'id':str(uuid.uuid1()),
                            'iterationType':'ALL', #ALL|RANGE|SPECIFIC
                            'testDataId':'Data Files/'+testDataInfo.get("name"),
                            'value':'',
                            })
                data.get("testCases").append( testCaseInfo )
        return (0, data)

    @classmethod
    def create_suite_xml( cls, baseInfo, caseList ):
        ret, suiteInfo = cls.get_case_data_vars( baseInfo, caseList )
        if ret != 0:
            logging.error( u'创建Test Suite失败[%s]'%(baseInfo.get('suiteName') ) )
        if suiteInfo:
            project_path = cls.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
            suiteName = os.path.join( project_path, "Test Suites", suiteInfo.get("suiteName")+".ts" )
            testCases = suiteInfo.get("testCases")
            try:
                #Test Suite描述
                root = ElementTree.Element("TestSuiteEntity")
                node_desc = ElementTree.SubElement( root, "description" )
                node_desc.text=suiteInfo.get('description')

                #Test Suite名称
                node_name = ElementTree.SubElement( root, "name" )
                node_name.text = suiteInfo.get('suiteName')

                node_tag = ElementTree.SubElement( root, 'tag' )

                #是否重复运行
                node_isRerun = ElementTree.SubElement( root, 'isRerun' )
                node_isRerun.text = suiteInfo.get('isRerun')
                
                #最后一次运行时间
                node_lastRun = ElementTree.SubElement( root, 'lastRun' )
                node_lastRun.text= suiteInfo.get('lastRun')

                node_mainRecipient = ElementTree.SubElement( root, 'mainRecipient')

                #重复运行次数
                node_numberOfRerun = ElementTree.SubElement( root, 'numberOfRerun' )
                node_numberOfRerun.text = suiteInfo.get('numberOfRerun')
                
                #加载超时时间
                node_pageLoadTimeout = ElementTree.SubElement( root, 'pageLoadTimeout' )
                node_pageLoadTimeout.text = suiteInfo.get('pageLoadTimeout')

                #使用默认加载超时时间
                node_pageLoadTimeoutDefault = ElementTree.SubElement( root, 'pageLoadTimeoutDefault' )
                node_pageLoadTimeoutDefault.text = suiteInfo.get('pageLoadTimeoutDefault')

                #只在有案例失败时重复运行
                node_rerunFailedTestCasesOnly = ElementTree.SubElement( root, 'rerunFailedTestCasesOnly' )
                node_rerunFailedTestCasesOnly.text = suiteInfo.get('rerunFailedTestCasesOnly')

                #TestSuite 全局ID
                node_testSuiteGuid = ElementTree.SubElement( root, 'testSuiteGuid' )
                node_testSuiteGuid.text = str(uuid.uuid1())


                for testCase in testCases:
                    #案例节点
                    node_testCase = ElementTree.SubElement( root, "testCaseLink" )

                    #案例的GUID
                    node_guid = ElementTree.SubElement( node_testCase, "guid" )
                    node_guid.text=str(uuid.uuid1())
                    #time.sleep(1)

                    #是否复用驱动
                    node_isReuseDriver = ElementTree.SubElement( node_testCase, "isReuseDriver" )
                    node_isReuseDriver.text=testCase.get('isReuseDriver')

                    #是否运行
                    node_isRun = ElementTree.SubElement( node_testCase, "isRun" )
                    node_isRun.text = testCase.get('isRun')
                    
                    #案例的ID
                    node_testCaseId = ElementTree.SubElement( node_testCase, "testCaseId" )
                    node_testCaseId.text = testCase.get('testCaseId')

                    testDataId=str(uuid.uuid1())

                    for testData in testCase.get("testDatas"):
                        #案例绑定的数据源
                        node_testData = ElementTree.SubElement( node_testCase, "testDataLink" )

                        #绑定类型 取值ONE|MANY
                        node_combinationType = ElementTree.SubElement( node_testData, "combinationType" )
                        node_combinationType.text = testData.get('combinationType')

                        #数据的UUID
                        node_id = ElementTree.SubElement( node_testData, "id" )
                        node_id.text = testDataId

                        #数据迭代入口
                        node_iterationEntity = ElementTree.SubElement( node_testData, 'iterationEntity' )

                        #数据迭代类型ALL|RANGE|SPECIFIC
                        node_iterationType = ElementTree.SubElement( node_iterationEntity, 'iterationType' )
                        node_iterationType.text = testData.get('iterationType')

                        #数据迭代值 对应迭代类型分别为 空|m-n|num
                        node_value = ElementTree.SubElement( node_iterationEntity, 'value' )
                        node_value.text = testData.get('value')

                        #数据源ID
                        node_testDataId = ElementTree.SubElement( node_testData, 'testDataId' )
                        node_testDataId.text = testData.get('testDataId')
                    for variable in testCase.get("variables"):
                        #变量节点
                        node_variableLink = ElementTree.SubElement( node_testCase, 'variableLink' )

                        #数据源ID
                        node_testDataLinkId = ElementTree.SubElement( node_variableLink, 'testDataLinkId' )
                        node_testDataLinkId.text = testDataId

                        #数据类型 取值 DATA_COLUMN|DATA_COLUMN_INDEX|DEFAULT|SCRIPT_VARIABLE
                        node_type = ElementTree.SubElement( node_variableLink, 'type' )
                        node_type.text= variable.get('type')

                        #数据值
                        node_value = ElementTree.SubElement( node_variableLink, 'value' )
                        node_value.text = variable.get('desc')+"("+variable.get('value')+")"

                        #变量UUID
                        node_variableId = ElementTree.SubElement( node_variableLink, 'variableId' )
                        node_variableId.text= variable.get('variableId')

                #el = ElementTree.ElementTree( root )
                doc = xml.dom.minidom.parseString(ElementTree.tostring(root))
                with open( suiteName, "w" ) as fp:
                    doc.writexml( fp, indent="\t", newl="\n", encoding="gbk" )
            except Exception as e:
                logging(u'创建测试套件错误[%s]'%(e.message))
                return (-1, u'生成TestSuite失败[%s]'%(e.message) )
        return (0, u'生成TestSuite[%s]成功'%(suiteName) )

    @classmethod
    def Executeable( cls, cmd ):
        #cmd命令是否可执行
        p = subprocess.Popen( 'where %s'%(cmd), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True )
        output,outerr = p.communicate()
        if p.returncode == 0:
            logging.info(output)
            return True
        else:
            logging.warn(output)
            return False

    @classmethod
    def GetDeviceInfo( cls, project_path, key=None ):
        #获取设备信息，返回字典
        for root, dirs, files in os.walk( project_path ):
            for item in files:
                if item == "com.kms.katalon.core.mobile.android.properties":
                    with open( os.path.join( root, item ) ) as fp:
                        kv = json.loads(fp.read())
                        if kv is not None:
                            skv = kv.get("ANDROID_DRIVER")
                            if skv is not None:
                                if key is not None:
                                    if skv.get(key) is not None:
                                        return {key:skv.get(key)}
                                else:
                                    return skv
                        else:
                            return None
        return None

    @classmethod
    def SetDeviceInfo( cls, project_path, kv ):
        #设置设备信息，修改Katalon默认运行设备
        for root, dirs, files in os.walk( project_path ):
            for item in files:
                if item == "com.kms.katalon.core.mobile.android.properties":
                    deviceInfo = None
                    with open( os.path.join( root, item ), 'r' ) as fp:
                        deviceInfo = json.loads(fp.read())
                    if deviceInfo:
                        dkv = deviceInfo.get("ANDROID_DRIVER")
                        if dkv:
                            for key in kv:
                                if key in dkv:
                                    dkv[key]=kv.get(key)
                                    deviceInfo["ANDROID_DRIVER"]=dkv
                            try:
                                with open( os.path.join( root, item ) , "w" ) as fpw:
                                    fpw.write(json.dumps(deviceInfo) )
                                    return (0, u'设置成功')
                            except Exception as e:
                                return (1, u'设备信息设备设置失败!' )
                        else:
                            return (1, u'请在Katalon开发工具中设置设备信息!')
                    else:
                        return (1, u'请在Katalon开发工具中设置设备信息!')
    @classmethod
    def ConfigRead( cls, filename, section, key=None ):
        #读取配置文件中的SESSION字典或某一KEY值得字典
        try:
            config = ConfigParser.ConfigParser()
            config.read( filename )
            if config.has_section(section):
                if key:
                    if config.has_option( section, key ):
                        return config.get( section, key )
                else:
                    d={}
                    for key, value in config.items( section ):
                        d[key]=value
                    return d
        except Exception as e:
            logging.error(u"读配置文件错误[%s]!"%(e.message))
            return None

    @classmethod
    def ConfigWrite( cls, filename, section, key, value ):
        try:
            config = ConfigParser.ConfigParser()
            config.read(filename)
            with open( filename, "w" ) as fp:
                if not config.has_section( section ):
                    config.add_section( section )
                config.set( section, key, value )
                config.write(fp)
        except Exception as e:
            logging.error(u'写配置文件错误[%s]'%(e.message))

if __name__ == '__main__':
    #print commonUtils.create_data_xml_2("loanApply", data=None)
    #commonUtils.update_test_case( "loanApply.tc", 2 )
    #print commonUtils.parse_test_case_2('loanApply.tc')
    #print commonUtils.create_suite_xml({"suiteName":"suite"})
    #suiteInfo = {"suiteName":r'E:\MBKAutoTest\Test Suites\todaySuite.ts', "name":"todaySuite", "description":u"今天的测试案例", "numberOfRerun":"0", "pageLoadTimeout":"30", "pageLoadTimeoutDefault":"true", "returnFailedTestCasesOnly":"false"}
    #print commonUtils.GetDeviceInfo("E:\MBKAutoTest")
    #print commonUtils.SetDeviceInfo("E:\MBKAutoTest", {"deviceId":"12345"})
    #colprint commonUtils.Executeable("katalon")
    pass
