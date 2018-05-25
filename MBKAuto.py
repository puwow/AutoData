#-*- coding:gbk -*-

import os
import re
import threading
import logging
import subprocess
import datetime
import random
import wx
import wx.aui
import xlrd
import wx.propgrid as wxpg
import wx.lib.mixins.listctrl as listmix
import wx.adv
from wx.lib.wordwrap import wordwrap
import wx.lib.intctrl
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger
from commonUtils import commonUtils
import AutoTask
import dbhash
import shelve
import pytz
import web

logging.basicConfig(filename="mbkauto.log", filemode="a", level=logging.INFO, format="%(asctime)s:%(levelname)s:%(funcName)s:%(lineno)d:%(message)s")
CONFIG_FILE = "config.ini"
she = shelve.open( "memo.she", writeback=True)
AutoTask.initDB()

class CheckEnvironThread( threading.Thread ):
    def __init__( self, log, window, pack_env = 'VIRT' ):
        threading.Thread.__init__( self )
        self.log = log
        self.window = window
        self.pack_env = pack_env
    def run( self ):
        self.window.btn_check.Enable(False)
        self.log.DeleteAllItems()
        flag=True
        run_flag = False
        run_flag,result = AutoTask.hasRunning( pack_env=self.pack_env )

        #校验包管理服务
        message=''
        imageIndex=0
        try:
            AutoTask.checkServer( timeout=5 )
        except Exception as e:
            flag=False
            imageIndex=1
        finally:
            self.window.writeRecord( u'校验版本宝服务', imageId=imageIndex)

        #校验SVN
        message=''
        imageIndex=0
        svn_url = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "svn_url" )
        try:
            AutoTask.checkSVN()
        except Exception as e:
            flag=False
            imageIndex=1
        finally:
            self.window.writeRecord( u'校验SVN服务', imageId=imageIndex )

        #校验Katalon
        katalonFlag = True
        message=''
        imageIndex=0
        katalonFlag = AutoTask.checkKatalon()
        if katalonFlag is False:
            flag=False
            imageIndex=1

        self.window.writeRecord( u'校验Katalon环境', imageId=imageIndex )

        #校验手机设备
        message=''
        imageIndex=0
        try:
            kv = AutoTask.checkDevice()
            if kv.get('deviceId') == '' or kv.get('deviceId') is None:
                flag=False
                imageIndex=1
        except Exception as e:
            imageIndex=1
            flag=False
        finally:
            self.window.writeRecord(u'校验手机设备连接', imageId=imageIndex )

        if (flag is True) and (run_flag is False):
            self.window.btn_start.Enable(True)
        self.window.btn_check.Enable(True)

class RunningThread( threading.Thread ):
    def __init__( self, packInfo, maxCase, deviceId, log, window ):
        threading.Thread.__init__( self )
        self.maxCase = maxCase
        self.deviceId = deviceId
        self.log = log
        self.window = window
        self.packInfo = packInfo

    def run( self ):
        #自动化测试项目路径
        project_path = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
        #存放随机案例数
        minCase=1
        self.maxCase=1
        count=0
        if self.maxCase <= minCase:
            count = random.randint( self.maxCase, minCase )
        else:
            count = random.randint( minCase, self.maxCase )

        #触发Slider控件事件，随机获得count个案例
        wx.PostEvent( self.window, wx.CommandEvent(wx.EVT_SCROLL_CHANGED.typeId, self.window.slider.GetId()) )

        if self.packInfo is None:
            self.window.writeRecord( u'没有需要测试的程序包' )
            return

        packName = self.packInfo.get('pack_name')
        packDesc = self.packInfo.get('pack_desc')

        idx = 0
        try:
            #更新包状态
            self.window.UpdatePackStatus( self.packInfo, status='1' )
            #刷新包队列
            self.window.ShowQueuePack( pack_env = self.window.ctrl_Env.GetClientData( self.window.ctrl_Env.GetSelection() ) )
            #显示正在测试中的包
            self.window.ShowRunningPack( pack_env = self.window.ctrl_Env.GetClientData( self.window.ctrl_Env.GetSelection() ) )
            #更新案例数据
            self.window.UpdateData()

            self.window.writeRecord( u'更新案例数据', imageId=0 )
            try:
                #生成测试套件
                suiteName = self.window.CreateTestSuite( packName=packName, packDesc = packDesc )
                self.window.writeRecord( u'生成测试套件', imageId=0 )
                try:
                    #同步运行设备
                    self.window.SyncDevice( project_path = project_path, deviceId=self.deviceId )
                    self.window.writeRecord( u'同步运行设备', imageId=0 )
                    self.window.autoRun( project_path, suiteName, self.deviceId )
                except Exception as e:
                    self.window.writeRecord( imageId=1, flag=1 )
                    raise Exception(e.message)
            except Exception as e:
                self.window.writeRecord( imageId=1, flag=1 )
                raise Exception(e.message)
        except Exception as e:
            self.window.writeRecord( imageId=1, flag=1 )
            self.window.UpdatePackStatus( self.packInfo, status='0' )
            raise Exception(e.message)


class TestDataList( wx.ListCtrl, listmix.ListCtrlAutoWidthMixin, listmix.TextEditMixin ):
    def __init__( self, parent, data=None, data2=None, id=wx.ID_ANY, style=wx.LC_REPORT|wx.LC_VRULES|wx.LC_HRULES, name="data" ):
        wx.ListCtrl.__init__( self, parent=parent, id=id, style=style, name=name )
        listmix.ListCtrlAutoWidthMixin.__init__( self )
        listmix.TextEditMixin.__init__( self )
        #案例信息
        self.data = data
        #案例数据
        self.data2 = data2
        #参数和数据列是否相等
        self.notEqual = False
        self.Bind( wx.EVT_CONTEXT_MENU, self.OnRightClick )
        self.selected = -1 #默认无数据
        #被选择的行
        self.ncolour = wx.Colour(0,0,0,255)
        self.gcolour = wx.Colour(35,142,35,255)
        self.InitColumn()
    def InitColumn( self ):
        n = 1
        if self.data:
            ret, test_case_info = commonUtils.parse_test_case_2( self.data.get('path') )
            if ret == 0:
                variables = test_case_info.get("variables")
                self.data['param_count']=len(variables)
                #变量数组
                i=0
                for var in variables:
                    #变量字典
                    if var[1] is not None:
                        self.InsertColumn( i, var[1]+"("+var[0]+")" )
                    else:
                        self.InsertColumn( i, var[0] )
                    i=i+1
                    reg = re.search(r'(\d).+(\d)', var[2])
                    if reg:
                         #数据在TestCase中绑定的行
                         n = int(reg.groups()[0])
                         self.selected = n
        self.InitData(n, test_case_info.get('name'))
    def isNotE( self ):
        return self.notEqual
    def InitData( self, n, name ):
        #初始化数据, n是从tc中取得的默认行数，正常情况比实际行数大1
        ccol = self.GetColumnCount()
        row = 0
        notEqual = False
        for record in self.data2[1:]:
            if len(record) != ccol:
                #如果从TC中获取的变量数与Excel中的数据列数不一致，丢弃多余数据
                logging.warn(u'测试案例[%s]数据与变量不一致'%(name))
                self.notEqual = True
                record=record[:ccol]
            col = 0
            for item in record:
                if col == 0:
                    try:
                        self.InsertItem( row, str(item) )
                    except Exception as e:
                        self.InsertItem( row, item )
                else:
                    try:
                        self.SetItem( row, col, str(item) )
                    except Exception as e:
                        self.SetItem( row, col, item )
                col=col+1
                if col == ccol:
                    break
            row=row+1
        if n>=1 and n-1 < len(self.data2[1:]) and self.notEqual is False:
            self.SetItemTextColour( n-1, self.gcolour)
    def GetData( self ):
        #提供数据给Excel表
        data=[]
        udata=[]
        for col in range( self.GetColumnCount() ):
            column = self.GetColumn( col )
            udata.append(column.GetText())
        data.append(udata)
        sel = -1
        while True:
            #行变量
            sel = self.GetNextItem( sel, wx.LIST_NEXT_ALL )
            if sel == -1:
                break
            udata=[]
            for col in range( self.GetColumnCount() ):
                udata.append( self.GetItemText( sel, col ) )
            data.append(udata)
        self.data['data']=data
        return self.data
    def OnRightClick( self, event ):
        if not hasattr( self, "addMenuId" ):
            self.addMenuId = wx.NewId()
            self.deleteMenuId = wx.NewId()
            self.defaultMenuId = wx.NewId()
            self.Bind( wx.EVT_MENU, self.OnAdd, id=self.addMenuId )
            self.Bind( wx.EVT_MENU, self.OnDelete, id=self.deleteMenuId )
            self.Bind( wx.EVT_MENU, self.OnSetDefault, id=self.defaultMenuId )
        menu = wx.Menu()
        addMenu = wx.MenuItem( menu, id=self.addMenuId, text=u"添加数据", helpString=u'添加一条数据' )
        deleteMenu = wx.MenuItem( menu, id=self.deleteMenuId, text=u"删除数据", helpString=u'删除一条数据')
        defaultMenu = wx.MenuItem( menu, id=self.defaultMenuId, text=u'设为默认值', helpString=u'设置该条记录为默认值')
        if os.path.exists( "images/add24.png" ):
            bmp = wx.Bitmap( "images/add24.png", wx.BITMAP_TYPE_PNG )
            addMenu.SetBitmap( bmp )
        if os.path.exists( "images/delete24.png" ):
            bmp = wx.Bitmap( "images/delete24.png", wx.BITMAP_TYPE_PNG )
            deleteMenu.SetBitmap( bmp )
        if os.path.exists( 'images/default24.png'):
            bmp = wx.Bitmap( "images/default24.png", wx.BITMAP_TYPE_PNG )
            defaultMenu.SetBitmap( bmp )
        menu.Append( addMenu )
        menu.Append( deleteMenu )
        menu.Append( defaultMenu )

        self.PopupMenu( menu )
        menu.Destroy()
    def OnAdd( self, event ):
        i = self.GetColumnCount()
        if i != 0:
            self.InsertItem( i, "" )
    def OnDelete( self, event ):
        sel = self.GetFirstSelected()
        if sel != -1:
            logging.info(u"删除记录[%s]"%(self.GetItemText( sel, 1 )))
            self.DeleteItem( sel )
    def OnSetDefault( self, event ):
        sel = -1
        while True:
            sel = self.GetNextItem( sel, wx.LIST_NEXT_ALL )
            if sel == -1:
                break
            self.SetItemTextColour(sel, self.ncolour )
        row = self.GetFirstSelected()
        #设置默认值，ListCtrl中的值从0起步，Excel表中的值从1起步
        if row <=0:
            self.selected = 1
        if row >=0:
            self.selected = row + 1
            self.SetItemTextColour(row, self.gcolour)
        logging.info(u'设置第[%d]条记录为默认执行数据!'%(self.selected))
    def GetDefaultData( self ):
        if self.selected <=0:
            return 1
        else:
            return self.selected

class TestCaseTree( wx.TreeCtrl ):
    def __init__( self, parent, id=wx.ID_ANY, style=wx.TR_HAS_BUTTONS, name="tree" ):
        wx.TreeCtrl.__init__( self, parent=parent, id=id, style=style, name="tree" )
        self.CreateImage()
        self.Bind( wx.EVT_TREE_SEL_CHANGED, self.OnSelChanged, self )
        self.Bind( wx.EVT_RIGHT_DOWN, self.OnRightClick )
        self.Bind( wx.EVT_LEFT_DCLICK, self.OnDBClick )
        wx.CallAfter(self.InitData)
        self.count = 0
    def InitData( self ):
        self.DeleteAllItems()
        try:
            dict_cases = commonUtils.find_test_cases()
            logging.info(dict_cases)
            root = self.AddRoot(u"项目案例集合", image=4, selImage=4, data=dict_cases)
            if dict_cases:
                count=0
                for key in dict_cases.keys():
                    logging.info(key)
                    desc=''
                    key_val=key
                    if she is not None:
                        #存储Key和取出Key值
                        if key not in she.keys():
                            she[str(key)]={'desc':'', 'childs':[]}
                        else:
                            desc=she[str(key)]['desc']
                    #二级节点
                    if desc != '':
                        key_val=key_val+"("+unicode(desc)+")"
                    node = self.AppendItem( root, key_val, image=0, selImage=0, data=dict_cases.get(key) )
                    for case in dict_cases.get(key):
                        if case.get('name') not in [ x[0] for x in she[str(key)].get('childs') ]:
                            she[str(key)].get('childs').append((case.get('name'),0))
                        else:
                            for x in she[str(key)].get('childs'):
                                if case.get('name') == x[0]:
                                    case['weight']=x[1]
                        #三级节点
                        index = 1
                        if case.get("desc") == u"未标记":
                            index=3
                        self.AppendItem( node, case.get("name")+"("+case.get("desc")+")", image=index, selImage=2, data=[case] )
                        count=count+1
                logging.info(u'共加载测试案例[%d]条!'%(count))
                self.count = count
                self.Expand( root )
        except Exception as e:
            logging.error(e)
    def GetCaseCount( self ):
        return self.count
    def OnSelChanged( self, event ):
        self.item = event.GetItem()
    def CreateImage(self):
        imgL = wx.ImageList( 16, 16 )
        images = ['leafs.png', 'leaf_blue.png', 'leaf_red.png', 'leaf_yellow.png', 'tree.png']
        for image in images:
            if os.path.exists( os.path.join("images", image ) ):
                bmp = wx.Bitmap( os.path.join( "images", image ), wx.BITMAP_TYPE_PNG )
                imgL.Add( bmp )
        self.AssignImageList( imgL )
    def OnDBClick( self, event ):
        root = self.GetRootItem()
        item = self.GetFocusedItem()
        data = self.GetItemData( item )
        if data and (not self.ItemHasChildren( item ) and item != root):
            if self.GetParent().FindWindowByName("list"):
                self.GetParent().FindWindowByName("list").InitData( data )
                self.GetParent().FindWindowByName("list").Refresh()
                logging.info(u'测试套ADD案例[%s]'%(self.GetItemText(item)))
        event.Skip()
    def OnRightClick( self, event ):
        pt = event.GetPosition()
        item,flags = self.HitTest(pt)
        if item:
            self.SelectItem( item )
        if not hasattr( self, "addMenu" ):
            self.addMenuId = wx.NewId()
            self.freshMenuId = wx.NewId()
            self.expandMenuId = wx.NewId()
            self.collapseMenuId = wx.NewId()

            self.Bind( wx.EVT_MENU, self.OnAdd, id=self.addMenuId )
            self.Bind( wx.EVT_MENU, self.OnFresh, id=self.freshMenuId )
            self.Bind( wx.EVT_MENU, self.OnExpand, id=self.expandMenuId )
            self.Bind( wx.EVT_MENU, self.OnCollapse, id=self.collapseMenuId )
        menu = wx.Menu()
        addMenu = wx.MenuItem( menu, id=self.addMenuId, text=u'添加Case', helpString=u'添加一条记录' )
        freshMenu = wx.MenuItem( menu, id=self.freshMenuId, text=u'刷新Case', helpString=u'刷新记录' )
        expandMenu = wx.MenuItem( menu, id=self.expandMenuId, text=u'全部展开', helpString=u'展开全部子项' )
        collapseMenu = wx.MenuItem( menu, id=self.collapseMenuId, text=u'全部折叠', helpString=u'折叠全部子项' )
        if os.path.exists( "images/link_add24.png" ):
            bmp = wx.Bitmap( "images/link_add24.png", wx.BITMAP_TYPE_PNG )
            addMenu.SetBitmap( bmp )
        if os.path.exists( "images/link_refresh24.png" ):
            bmp = wx.Bitmap( "images/link_refresh24.png", wx.BITMAP_TYPE_PNG )
            freshMenu.SetBitmap( bmp )
        if os.path.exists( "images/expand24.png" ):
            bmp = wx.Bitmap( "images/expand24.png", wx.BITMAP_TYPE_PNG )
            expandMenu.SetBitmap( bmp )
        if os.path.exists( "images/collapse24.png" ):
            bmp = wx.Bitmap( "images/collapse24.png", wx.BITMAP_TYPE_PNG )
            collapseMenu.SetBitmap( bmp )

        menu.Append( addMenu )
        menu.Append( freshMenu )
        menu.Append( expandMenu )
        menu.Append( collapseMenu )

        self.PopupMenu( menu )
        menu.Destroy()
    def OnAdd( self, event ):
        sel = self.GetFocusedItem()
        data = []
        if sel:
            if self.GetRootItem() == sel:
                items = self.GetItemData( sel )
                for key in items:
                    for item in items.get(key):
                        data.append( item )
            else:
                data = self.GetItemData( sel )
        if self.GetParent().FindWindowByName("list"):
            self.GetParent().FindWindowByName("list").InitData( data )
            self.GetParent().FindWindowByName("list").Refresh()

    def OnFresh( self, event ):
        self.InitData()
    def OnExpand( self, event ):
        self.ExpandAll()
    def OnCollapse( self, event ):
        item, cookie =  self.GetFirstChild( self.GetRootItem() )
        while item:
            self.Collapse( item )
            item, cookie = self.GetNextChild( self.GetRootItem(), cookie )

class TestCaseList( wx.ListCtrl ):
    def __init__( self, parent, id=wx.ID_ANY, style=wx.LC_REPORT|wx.LC_VRULES|wx.LC_HRULES, name="list" ):
        wx.ListCtrl.__init__( self, parent=parent, id=id, style=style, name="list" )

        self.autoAdd=[]
        self.userAdd=[]
        self.InsertColumn( 0, u"案例编号" )
        self.InsertColumn( 1, u"案例名称", width=160 )
        self.InsertColumn( 2, u"案例描述", width=240 )
        self.InsertColumn( 3, u"入参个数", width=80 )
        self.InsertColumn( 4, u"数据量", width=80)
        self.InsertColumn( 5, u"权重", width=80)
        self.InsertColumn( 6, u"脚本路径", width=-1)


        self.Bind( wx.EVT_CONTEXT_MENU, self.OnRightClick )
        self.Bind( wx.EVT_LIST_ITEM_SELECTED, self.OnSelect )
        wx.CallAfter( self.InitData )

    def OnSelect( self, event ):
        sel = self.GetFirstSelected()
        nb = wx.GetApp().GetTopWindow().GetNoteBook()
        if sel != -1:
            title = self.GetItemText( sel, 1 )
            for i in range( nb.GetPageCount() ):
                if nb.GetPageText(i) == title:
                    nb.SetSelection(i)

    def InitData( self, data=None, auto=False ):
        busy = wx.BusyInfo(u"数据加载，请稍后...", parent=self)
        nb = self.GetParent().GetParent().GetNoteBook()
        if data is None:
            data=[]
        sel = -1
        nb.Freeze()
        self.DeleteAllItems()
        nb.DeleteAllPages()
        self.autoAdd = []
        allAdd=self.userAdd


        if auto is False:
            #手动添加
            for item in data:
                if item not in self.userAdd:
                    self.userAdd.append( item )
        else:
            #自动添加
            for item in data:
                if item  not in self.autoAdd:
                    if item not in allAdd:
                        self.autoAdd.append( item )
        allAdd = allAdd + self.autoAdd

        i=self.GetItemCount()
        for item in  sorted( allAdd, cmp=lambda x,y:cmp(x.get('weight'),y.get('weight')), reverse=True ):
            #权重转换
            weight_info = ''
            weight = item.get('weight')
            if weight == 0:
                weight_info = u'不重要'
            elif weight == 1:
                weight_info = u'一般'
            elif weight == 2:
                weight_info = u'中等'
            elif weight == 3:
                weight_info = u'重要'
            self.InsertItem( i, str(i+1) )
            self.SetItem( i, 1, item.get("name") )
            self.SetItem( i, 2, item.get("desc") )

            if item.get("desc") == u"未标记":
                self.SetItemBackgroundColour(i, "yellow")

            #读取当前Sheet页的数据
            rdata = commonUtils.read_excel( item.get("name") )

            #数据量
            data_count=0
            if len(rdata) == 1 or len(rdata) == 0:
                data_count=0
            else:
                data_count=len(rdata)-1



            #入参个数
            param_count = item.get('param_count', 0)

            self.SetItem( i, 3, str(param_count) )
            self.SetItem( i, 4, str(data_count) )
            self.SetItem( i, 5, weight_info )
            self.SetItem( i, 6, item.get('path') )

            td = None
            td = TestDataList( nb, item, data2=rdata )
            if td.isNotE() is True:
                self.SetItemTextColour(i, "red")
            i=i+1

            if td is not None:
                nb.AddPage( td, item.get("name"), select=False, imageId=0 )
                nb.SetPageToolTip( nb.GetPageIndex(td), item.get("desc") )
            logging.info(u"测试套ADD案例[%s(%s)]"%(item.get("name"),item.get("desc")))
        nb.Thaw()
        del busy
    def OnRightClick( self, event ):
        if not hasattr( self, "deleteMenu" ):
            self.deleteMenuId = wx.NewId()
            self.tempMenuId = wx.NewId()
            self.clearMenuId = wx.NewId()
            self.suiteMenuId = wx.NewId()

            self.Bind( wx.EVT_MENU, self.OnDelete, id=self.deleteMenuId )
            self.Bind( wx.EVT_MENU, self.OnTemp, id=self.tempMenuId )
            self.Bind( wx.EVT_MENU, self.OnClear, id=self.clearMenuId )
            self.Bind( wx.EVT_MENU, self.OnTestSuite, id=self.suiteMenuId )
        menu = wx.Menu()
        deleteMenu = wx.MenuItem( menu, id=self.deleteMenuId, text=u'删除Case', helpString=u'删除一条记录' )
        clearMenu = wx.MenuItem( menu, id=self.clearMenuId, text=u'清空记录', helpString=u'清空记录' )
        tempMenu = wx.MenuItem( menu, id=self.tempMenuId, text=u'更新数据模板', helpString=u'生成被选择案例的测试数据模板' )
        suiteMenu = wx.MenuItem( menu, id=self.suiteMenuId, text=u'生成TestSuite', helpString=u'根据列表生成TestSuite' )
        if os.path.exists( "images/link_delete24.png" ):
            bmp = wx.Bitmap( "images/link_delete24.png", wx.BITMAP_TYPE_PNG )
            deleteMenu.SetBitmap( bmp )
        if os.path.exists( "images/excel24.png" ):
            bmp = wx.Bitmap( "images/excel24.png", wx.BITMAP_TYPE_PNG )
            tempMenu.SetBitmap( bmp )
        if os.path.exists( "images/clean24.png" ):
            bmp = wx.Bitmap( "images/clean24.png", wx.BITMAP_TYPE_PNG )
            clearMenu.SetBitmap( bmp )
        if os.path.exists( "images/suite24.png" ):
            bmp = wx.Bitmap( "images/suite24.png", wx.BITMAP_TYPE_PNG )
            suiteMenu.SetBitmap( bmp )

        menu.Append( deleteMenu )
        menu.Append( clearMenu )
        menu.AppendSeparator()
        menu.Append( tempMenu )
        menu.Append( suiteMenu )
        self.PopupMenu( menu )
        menu.Destroy()
    def OnDelete( self, event ):
        nb = wx.GetApp().GetTopWindow().GetNoteBook()
        sel = self.GetFirstSelected()
        if sel != -1:
            case_name = self.GetItemText( sel, 1 )
            for i in range( nb.GetPageCount() ):
                if nb.GetPageText(i) == case_name:
                    nb.DeletePage( i )
            logging.info(u'测试套DELETE案例[%s(%s)]'%(case_name, self.GetItemText(sel, 2)))
            self.DeleteItem( sel )
            for item in self.userAdd:
                if item.get('name') == case_name:
                    self.userAdd.remove( item )
            for item in self.autoAdd:
                if item.get('name') == case_name:
                    self.userAdd.remove( item )
    def OnTemp( self, event, auto=False ):
        sel = -1
        nb = wx.GetApp().GetTopWindow().GetNoteBook()
        while True:
            sel = self.GetNextItem( sel, wx.LIST_NEXT_ALL )
            if sel == -1:
                break
            sheet_name =  self.GetItemText( sel, 1 )
            data = None
            for i in range( nb.GetPageCount() ):
                if nb.GetPageText(i) == sheet_name:
                    nb.SetSelection(i)
            if nb:
                page = nb.GetCurrentPage()
                if page:
                    data = page.GetData()
                    cols = [page.GetColumn(x).GetText() for x in range(page.GetColumnCount())]
                    row = page.GetDefaultData()
                    commonUtils.update_excel( sheet_name, data=data )
                    commonUtils.create_data_xml_2( sheet_name )
                    commonUtils.update_test_case( data.get('path'), row )
        if auto is False:
            wx.MessageBox(u"数据更新完成")
    def OnClear( self, event ):
        nb = wx.GetApp().GetTopWindow().GetNoteBook()
        self.DeleteAllItems()
        nb.DeleteAllPages()
        self.userAdd = []
        self.autoAdd = []
    def OnTestSuite( self, event ):
        sel = -1
        caseList = []
        while True:
            sel = self.GetNextItem( sel, wx.LIST_NEXT_ALL )
            if sel == -1:
                break
            caseList.append(self.GetItemText( sel, 6 ))

        if len(caseList) == 0:
            wx.MessageBox( u'请先添加测试案例' )
            return
        self.Freeze()
        try:
            parent = wx.GetApp().GetTopWindow()
            dialog = TestSuiteDialog(parent, caseList=caseList)
            dialog.CenterOnParent()
            self.Thaw()
            dialog.ShowModal()
            dialog.Destroy()
        except Exception as e:
            self.Thaw()
            raise Exception(e.message)

class TestCaseSetting( wx.Panel ):
    def __init__( self, parent, filename, sheetname, id=wx.ID_ANY ):
        wx.Panel.__init__( self, parent=parent, id=id )

        self.grid = XG.XLSGrid( self )
        busy = wx.BusyInfo(u"正在读取数据文件，请等候...", parent=parent)
        try:
            book = xlrd.open_workbook( filename, formatting_info=1 )
            sheet = book.sheet_by_name( sheetname )
            rows, cols = sheet.nrows, sheet.ncols
            comments, texts = XG.ReadExcelCOM( filename, sheetname, rows, cols )
        finally:
            del busy
        self.grid.Show()
        self.grid.PopulateGrid( book, sheet, texts, comments )
class SettingDialog( wx.Dialog ):
    def __init__( self, parent=None, title=u'对话框', size=(600,480), style=wx.DEFAULT_DIALOG_STYLE ):
        wx.Dialog.__init__( self, parent=parent, title=title, size=size, style=style )

        panel = wx.Panel( self, -1 )
        self.pg = wxpg.PropertyGridManager( panel, style=wxpg.PG_SPLITTER_AUTO_CENTER|wxpg.PG_AUTO_SORT|wxpg.PG_TOOLBAR )
        self.pg.SetExtraStyle( wxpg.PG_EX_HELP_AS_TOOLTIPS )
        self.pg.AddPage(u"手机银行自动化测试全局设置")
        self.pg.Append( wxpg.DirProperty( u"项目目录", label="project_path", value="" ) )
        self.pg.Append( wxpg.DirProperty( u"数据存储路径", label="data_path", value="" ) )
        self.pg.Append( wxpg.BoolProperty( u'数据单文件', name="single_excel", value=True ) )
        self.pg.Append( wxpg.StringProperty( u'数据文件名', name="data_name", value=u"TestData.xls" ) )
        self.pg.Append( wxpg.FileProperty( u'Katalon命令', name="katalon_exe", value="" ) )
        self.pg.Append( wxpg.StringProperty( u'项目SNV地址', name="svn_url", value="svn://14.16.18.5/MBPSYS/trunk/MBKAutoTest" ) )
        self.pg.Append( wxpg.StringProperty( u'测试包服务地址', name="pack_url", value="http://13.239.21.170:8080/" ) )
        self.pg.Append( wxpg.FileProperty( u'adb命令', name="adb_exe", value="" ) )

        vbox = wx.BoxSizer( wx.VERTICAL )
        vbox.Add( self.pg, 1, wx.EXPAND|wx.ALL, 3 )

        btn_save = wx.Button( panel, -1, u'保存' )
        btn_close = wx.Button( panel, -1, u'关闭' )
        if os.path.exists( "images/save24.png" ):
            bmp = wx.Bitmap( "images/save24.png", wx.BITMAP_TYPE_PNG )
            btn_save.SetBitmap( bmp )
        if os.path.exists( "images/close24.png" ):
            bmp = wx.Bitmap( "images/close24.png", wx.BITMAP_TYPE_PNG )
            btn_close.SetBitmap( bmp )

        hbox = wx.BoxSizer( wx.HORIZONTAL)
        hbox.Add( btn_save, 0, wx.ALL|wx.ALIGN_CENTER, 3 )
        hbox.Add( btn_close, 0, wx.ALL|wx.ALIGN_CENTER, 3 )
        vbox.Add( hbox, 0, wx.ALL|wx.ALIGN_CENTER, 3  )
        panel.SetSizer( vbox )

        self.Bind( wx.EVT_BUTTON, self.OnSave, btn_save )
        self.Bind( wx.EVT_BUTTON, self.OnClose, btn_close )
        wx.CallAfter(self.InitData)
    def InitData( self ):
        kv = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST" )
        self.pg.SetValues( kv )
    def OnSave( self, event ):
        d = self.pg.GetValues(inc_attributes=False)
        for key in d:
            commonUtils.ConfigWrite( CONFIG_FILE, "MBKAUTOTEST", key, d[key] )
            logging.info(u'更改全局配置[%s:%s]'%(key, d[key]))
        wx.MessageBox(u"设置成功")
        if self.GetParent():
            if self.GetParent().FindWindowByName("list"):
                self.GetParent().FindWindowByName("list").InitData([])
            if self.GetParent().FindWindowByName("tree"):
                self.GetParent().FindWindowByName("tree").InitData()
        self.Destroy()
    def OnClose( self, event ):
        self.Destroy()

class AutoFrame( wx.Frame ):
    def __init__( self, parent=None, title=u"手机银行自动化测试数据管理", size=(800,600), style=wx.DEFAULT_FRAME_STYLE ):
        wx.Frame.__init__( self, parent=parent, title=title, size=size , style=style)

        mainPanel = wx.Panel( self, id=wx.ID_ANY )
        self.tree = TestCaseTree( parent=mainPanel, id=wx.ID_ANY, name="tree")

        self.list = TestCaseList( parent=mainPanel, id=wx.ID_ANY, name="list" )

        self.nb = wx.aui.AuiNotebook( parent=mainPanel, id=wx.ID_ANY, style=wx.aui.AUI_NB_TOP|wx.aui.AUI_NB_TAB_SPLIT|\
                wx.aui.AUI_NB_SCROLL_BUTTONS|wx.aui.AUI_NB_CLOSE_ON_ACTIVE_TAB|wx.aui.AUI_NB_MIDDLE_CLICK_CLOSE)

        self.autoPanel = AutoTestPanel( parent=mainPanel, id=wx.ID_ANY, name="set" )

        nb_images = wx.ImageList( 16, 16 )
        if os.path.exists("images/data.png"):
            bmp = wx.Bitmap( "images/data.png", wx.BITMAP_TYPE_PNG )
            nb_images.Add( bmp )
        self.nb.AssignImageList( nb_images )

        self.mgr = wx.aui.AuiManager()
        self.mgr.SetManagedWindow( mainPanel )
        self.SetMinSize( (800, 600) )

        self.mgr.AddPane( self.tree, wx.aui.AuiPaneInfo().Left().
                BestSize(-1,-1).
                MinSize(240,-1).
                Floatable(False).
                Caption(u'测试案例(Test Case)').
                CloseButton(False).
                PaneBorder(False).
                Name("TestCase"))
        self.mgr.AddPane( self.autoPanel, wx.aui.AuiPaneInfo().Right().
                BestSize(600,-1).
                MinSize(600, -1).
                Floatable(False).
                Caption(u'自动化设置').
                CloseButton(False).
                PaneBorder(False).
                Name("AutoSet"))
        self.mgr.AddPane( self.list, wx.aui.AuiPaneInfo().Center().
                MinimizeButton().
                MaximizeButton().
                PinButton().
                BestSize(-1,-1).
                MinSize(500,-1).
                Floatable(False).
                FloatingSize(500,160).
                Caption(u'测试套件(Test Suite)').
                CloseButton(False).
                PaneBorder(False).
                Name('TestCaseSetting'))
        self.mgr.AddPane( self.nb, wx.aui.AuiPaneInfo().Bottom().
                MinimizeButton().
                MaximizeButton().
                PinButton().
                BestSize(-1,400).
                MinSize(-1, 160).
                Floatable(False).
                FloatingSize(500, 240).
                Caption(u'测试数据(Test Data)').
                PaneBorder(False).
                CloseButton(False).
                Name('TestData'))
        self.mgr.Update()

        self.CreateMenuBar()
        self.Bind( wx.EVT_CLOSE, self.OnClose )
        #这个厉害了
        if os.path.exists("MBKAuto.exe"):
            self.SetIcon(wx.Icon("MBKAuto.exe", wx.BITMAP_TYPE_ICO) )
        logging.info(u"应用启动完成...")
    def OnClose( self, event ):
        logging.info(u"应用关闭...")
        if she:
            she.close()
        self.mgr.UnInit()
        self.Destroy()
    def CreateMenuBar( self ):
        menuBar = wx.MenuBar()

        menuGlobal = wx.Menu()
        menuSetting = wx.MenuItem( menuGlobal, id=wx.ID_ANY, text=u'设置(&S)', helpString=u'设置全局变量' )
        if os.path.exists( "images/setting24.png" ):
            bmp = wx.Bitmap( "images/setting24.png", wx.BITMAP_TYPE_PNG )
            menuSetting.SetBitmap( bmp )
        menuGlobal.Append( menuSetting )

        '''
        menuAuto = wx.MenuItem( menuGlobal, id=wx.ID_ANY, text=u'自动化设置(&T)', helpString=u'自动化设置')
        if os.path.exists( "images/auto24.png"):
            bmp = wx.Bitmap( "images/auto24.png", wx.BITMAP_TYPE_PNG)
            menuAuto.SetBitmap( bmp )
        menuGlobal.Append(menuAuto)
        '''

        menuHelp = wx.Menu()
        menuAbout = wx.MenuItem( menuHelp, id=wx.ID_ANY, text=u'关于(&A)', helpString=u'关于信息' )
        if os.path.exists( "images/about24.png" ):
            bmp = wx.Bitmap( "images/about24.png", wx.BITMAP_TYPE_PNG )
            menuAbout.SetBitmap( bmp )
        menuHelp.Append( menuAbout )

        menuMemo = wx.MenuItem(menuHelp, id=wx.ID_ANY, text=u'备注(&M)', helpString=u'备注信息' )
        if os.path.exists("images/memo24.png"):
            bmp = wx.Bitmap( "images/memo24.png")
            menuMemo.SetBitmap( bmp )
        menuHelp.Append( menuMemo )

        menuBar.Append( menuGlobal, u"全局(&G)" )
        menuBar.Append( menuHelp, u'帮助(&H)' )
        
        self.SetMenuBar( menuBar )
        self.Bind( wx.EVT_MENU, self.OnSetting, menuSetting )
        self.Bind( wx.EVT_MENU, self.OnAbout, menuAbout )
        self.Bind( wx.EVT_MENU, self.OnMemo, menuMemo )
    def OnSetting( self, event ):
        dialog = SettingDialog( parent=self, title=u'变量设置', size=(int(self.GetClientSize().width*0.65), int(self.GetClientSize().height*0.65)) )
        dialog.CenterOnParent()
        dialog.ShowModal()
        dialog.Destroy()
    def OnAbout( self, event ):
        info = wx.adv.AboutDialogInfo()
        info.Name=u"手机银行自动化测试数据管理工具"
        info.Version=u"1.1.0"
        info.Copyright=u"(c) 2018-2020 任何代码和程序的使用权限"
        info.Description=wordwrap(
                u"本工具是依托于Katalon软件测试项目结构，为了便于测试数据的管理与案例绑定而进行的二次定制开发，软件设计主要基于数据驱动测试的想法，目的在于解决测试数据的有效管理和减少人为干预案例的问题。\n\n本软件基于wxPython GUI库开发", 300, wx.ClientDC(self))
        
        info.Developers=['QiaoGS',]
        if os.path.exists("MBKAuto.exe"):
            info.SetIcon( wx.Icon( "MBKAuto.exe", wx.BITMAP_TYPE_ICO ) )
        wx.adv.AboutBox( info )
    def OnMemo( self, event ):
        dialog = MemoDialog( self, -1 )
        dialog.CenterOnParent()
        dialog.ShowModal()
    def GetNoteBook( self ):
        return self.nb
    def GetListCtrl( self ):
        return self.list
    def GetTreeCtrl( self ):
        return self.tree

class MemoDialog( wx.Dialog ):
    def __init__( self, parent, id=wx.ID_ANY, title=u'信息备注', size=(480,240) ):
        wx.Dialog.__init__( self, parent=parent, id=id, title=title, size=size )
        panel = wx.Panel( self, -1 )

        st_module = wx.StaticText( panel, -1, label=u'案例模块' )
        st_desc = wx.StaticText( panel, -1, label=u'模块备注' )
        st_cases = wx.StaticText( panel, -1, label=u'案例名称' )
        st_weight = wx.StaticText( panel, -1, label=u'权重' )
        self.ct_module = ct_module = wx.ComboBox( panel, -1 )
        self.ct_cases = ct_cases = wx.ComboBox( panel, -1 )
        self.ct_desc = ct_desc = wx.TextCtrl( panel, -1 )
        self.ct_weight = ct_weight = wx.Choice( panel, -1, choices=[u'不重要',u'一般',u'中等',u'重要'] )
        save = wx.Button( panel, id=wx.ID_ANY, label=u'保存')

        self.Bind( wx.EVT_COMBOBOX, self.OnChange, self.ct_module )
        self.Bind( wx.EVT_BUTTON, self.OnSave, save )
        self.Bind( wx.EVT_CLOSE, self.OnClose )
        self.Bind( wx.EVT_COMBOBOX, self.OnWeight, self.ct_cases )

        vbox = wx.BoxSizer( wx.VERTICAL )
        flex = wx.FlexGridSizer( 4, 2, 5, 5 )
        flex.AddMany([
            (st_module),
            (ct_module, 1, wx.EXPAND),
            (st_desc),
            (ct_desc, 1, wx.EXPAND),
            (st_cases),
            (ct_cases, 1, wx.EXPAND),
            (st_weight),
            (ct_weight, 1, wx.EXPAND)
            ])
        flex.AddGrowableCol( 1, 1 )
        vbox.Add( flex, 0, wx.EXPAND|wx.ALL, 10 )
        vbox.Add( save, 0, wx.ALL|wx.ALIGN_CENTER, 5 )
        panel.SetSizer( vbox )
        wx.CallAfter( self.initData )
    def OnWeight( self, event ):
        case_name = self.ct_cases.GetValue()
        module_name = self.ct_module.GetValue()
        for item in she[str(module_name)].get('childs'):
            if case_name == item[0]:
                self.ct_weight.SetSelection( item[1] )
    def OnSave( self, event ):
        idx = self.ct_weight.GetSelection()
        case = self.ct_cases.GetValue()
        if self.ct_module.GetValue() != '':
            if she is not None:
                try:
                    key = self.ct_module.GetValue()
                    she[str(key)]['desc']=str(self.ct_desc.GetValue())
                    for item in she[str(key)].get('childs'):
                        if case == item[0]:
                            she[str(key)].get('childs').remove(item)
                            she[str(key)].get('childs').append((case, idx))
                    wx.MessageBox(u'保存成功!')
                except Exception as e:
                    wx.MessageBox(u'保存失败[%s]'%(e.message))
    def initData( self ):
        if she is not None:
            try:
                for key in she:
                    self.ct_module.Append(key)
            except Exception as e:
                logging.error(e.message)
    def OnChange( self, event ):
        self.ct_cases.Clear()
        if she is not None:
            key = self.ct_module.GetValue()
            try:
                self.ct_desc.SetValue( unicode(she[str(key)].get('desc','')) )
                for item in she[str(key)].get('childs'):
                    self.ct_cases.Append( item[0] )
                    self.ct_weight.SetSelection( item[1] )
            except Exception as e:
                logging.error(e.message)
    def OnClose( self, event ):
        wlist = wx.GetApp().GetTopWindow().FindWindowByName("list")
        wtree = wx.GetApp().GetTopWindow().FindWindowByName("tree")
        if wlist:
            wlist.InitData([])
        if wtree:
            wtree.InitData()
        self.Destroy()


class TestSuiteDialog( wx.Dialog ):
    def __init__( self, parent, caseList=None, id=wx.ID_ANY, title=u'TestSuite组装', size=(540,420) ):
        wx.Dialog.__init__( self, parent=parent, id=id, title=title, size=size )

        self.caseList = caseList
        panel = wx.Panel( self, -1 )

        tz = pytz.timezone('Asia/Shanghai')
        now = datetime.datetime.now(tz)
        sb = wx.StaticBitmap( panel, id=wx.ID_ANY )
        tsName = wx.StaticText( panel, id=wx.ID_ANY, label=u'测试套件名称' )
        self.tc_tsName = wx.TextCtrl( panel, id=wx.ID_ANY, value=now.strftime('%Y%m%d') )
        tsDesc = wx.StaticText( panel, id=wx.ID_ANY, label=u'测试套件描述' )
        self.tc_tsDesc = wx.TextCtrl( panel, id=wx.ID_ANY, style=wx.TE_MULTILINE )
        tsRuntimes = wx.StaticText( panel, id=wx.ID_ANY, label=u'运行次数' )
        self.ic_tsRuntimes = wx.lib.intctrl.IntCtrl( panel, value=0, min=0 )
        tsTimeout = wx.StaticText( panel, id=wx.ID_ANY, label=u'超时时间' )
        self.ic_tsTimeout = wx.lib.intctrl.IntCtrl( panel, id=wx.ID_ANY, value=30, min=30 )
        tsDefaultTimeout = wx.StaticText( panel, id=wx.ID_ANY, label=u'使用默认超时时间')
        self.cb_tsDefaultTimeout = wx.CheckBox( panel, id=wx.ID_ANY )
        tsRerun = wx.StaticText( panel, id=wx.ID_ANY, label=u'仅重新运行错误案例' )
        self.cb_tsRerun = wx.CheckBox( panel, id=wx.ID_ANY )
        self.cb_tsDefaultTimeout.SetValue(True)

        self.tc_tsDesc.SetValue(now.isoformat())
        btn_run = wx.Button( panel, id=wx.ID_ANY, label=u"运行" )
        btn_save = wx.Button( panel, id=wx.ID_ANY, label=u'创建' )
        if os.path.exists("images/create32.png"):
            bmp = wx.Bitmap( "images/create32.png", wx.BITMAP_TYPE_PNG )
            btn_save.SetBitmap( bmp )
        if os.path.exists("images/run32.png"):
            bmp = wx.Bitmap( "images/run32.png", wx.BITMAP_TYPE_PNG )
            btn_run.SetBitmap( bmp )

        kata_device_label = wx.StaticText( panel, id=wx.ID_ANY, label=u'Katalon默认设备' )
        sys_device_label = wx.StaticText( panel, id=wx.ID_ANY, label=u'当前连接设备' )
        self.kata_device = wx.ComboBox( panel, id=wx.ID_ANY )
        self.sys_device = wx.ComboBox( panel, id=wx.ID_ANY )
        sync = wx.Button( panel, id=wx.ID_ANY, label=u'同步' )

        grid = wx.GridSizer( 1, 5, 5, 5 )
        grid.AddMany([(kata_device_label, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 3),\
                (self.kata_device, 0, wx.EXPAND|wx.ALIGN_CENTER_VERTICAL),\
                (sys_device_label, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL, 3 ),\
                (self.sys_device, 0, wx.EXPAND|wx.ALIGN_CENTER_VERTICAL, 3),\
                (sync, 0, wx.ALIGN_CENTER_VERTICAL, 3)
                ])

        hbox = wx.BoxSizer( wx.HORIZONTAL )
        hbox.Add( btn_save, 0, wx.ALIGN_CENTER|wx.RIGHT, 10 )
        hbox.Add( btn_run, 0, wx.ALIGN_CENTER ) 
        vbox = wx.BoxSizer( wx.VERTICAL )
        flex = wx.FlexGridSizer( 6, 2, 10, 10 )
        flex.AddMany([
                (tsName, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL),(self.tc_tsName, 0, wx.EXPAND),\
                (tsDesc, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL), (self.tc_tsDesc, 0, wx.EXPAND),\
                (tsRuntimes, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL), (self.ic_tsRuntimes, 0),\
                (tsTimeout, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL), (self.ic_tsTimeout, 0 ),\
                (tsDefaultTimeout, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL), (self.cb_tsDefaultTimeout, 0),\
                (tsRerun, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL), (self.cb_tsRerun, 0 )
                ])

        vbox2 = wx.BoxSizer( wx.VERTICAL )
        flex.AddGrowableCol(1,0)
        vbox.Add(flex, 3, wx.EXPAND|wx.ALL, 5 )
        vbox.Add( grid, 0, wx.ALL|wx.EXPAND|wx.BOTTOM, 10 )
        vbox2.Add( hbox, 0, wx.ALL|wx.ALIGN_CENTER,  3 )
        vbox.Add( vbox2, 1, wx.EXPAND )

        panel.SetSizer( vbox )

        self.Bind( wx.EVT_BUTTON, self.OnRun, btn_run )
        self.Bind( wx.EVT_BUTTON, self.OnSave, btn_save )
        self.Bind( wx.EVT_BUTTON, self.OnSync, sync )

        self.SetMinSize((540,420))
        self.InitDevice()
    def InitDevice( self ):
        self.kata_device.Clear()
        self.sys_device.Clear()
        projectPath = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
        if projectPath is not None:
            deviceId = commonUtils.GetDeviceInfo(projectPath, "deviceId")
            if deviceId is not None and type(deviceId) is dict:
                self.kata_device.Append( deviceId.get('deviceId') )
                self.kata_device.SetSelection(0)
        ret = commonUtils.Executeable("adb")
        if ret is True:
            p = subprocess.Popen( "adb devices", stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True, universal_newlines=True )
            output, outerr = p.communicate()
            if p.returncode == 0:
                logging.info(output)
                result = output.split('\n')
                for index in range(len(result)):
                    if index != 0:
                        deviceInfo = result[index].split('\t')
                        if len(deviceInfo)>0:
                            self.sys_device.Append( deviceInfo[0] )
                            self.sys_device.SetSelection(0)
            else:
                logging.error(output)
    def OnSync( self, event ):
        self.InitDevice()
        if self.sys_device.GetValue() == '':
            wx.MessageBox(u"未探测到手机连接")
            return
        if self.sys_device.GetValue() != self.kata_device.GetValue():
            dialog = wx.MessageDialog( self, message=u"是否将当前连接手机与Katalon默认手机同步?", caption=u'同步信息', style=wx.OK|wx.CANCEL )
            ret = dialog.ShowModal()
            if ret == wx.ID_OK:
                projectPath = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
                kv={"deviceId":self.sys_device.GetValue()}
                commonUtils.SetDeviceInfo(projectPath, kv )
                self.InitDevice()
    def OnRun( self, event ):
        if self.sys_device.GetValue() != self.kata_device.GetValue():
            wx.MessageBox(u"请先同步设备信息!")
            return
        ret = commonUtils.Executeable("katalon")
        if ret == False:
            wx.MessageBox(u"请将命令[katalon]命令添加到环境变量PATH中")
            return
        projectPath = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
        projectFile = os.path.join(commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" ), "MBKAutoTest.prj")
        suiteName = self.tc_tsName.GetValue()
        #if not re.match("r.*.ts", suiteName):
        #    suiteName = suiteName
        suitePath = os.path.join( "Test Suites", suiteName ).replace("\\","/")
        deviceId=self.kata_device.GetValue()
        cmd = 'katalon -runMode=console -consoleLog -noExit -projectPath="%s" -statusDelay=95000 -retry=0 -testSuitePath="%s" -deviceId="%s" -browserType="Android"'%(projectFile,suitePath, deviceId)
        p = subprocess.Popen( cmd, stdout=subprocess.PIPE, stderr = subprocess.STDOUT, shell=True )
        output,outerr = p.communicate()
        if p.returncode != 0:
            wx.MessageBox(u"执行错误!")
            logging.info(u"执行错误[%s]!"%(output) )
        logging.info(u"运行TestSuite[%s]"%(self.tc_tsName.GetValue()))
    def OnSave( self, event ):
        suiteInfo = {}
        suiteName = self.tc_tsName.GetValue()
        description = self.tc_tsDesc.GetValue()
        numberOfRerun = self.ic_tsRuntimes.GetValue()
        pageLoadTimeout = self.ic_tsTimeout.GetValue()
        pageLoadTimeoutDefault = self.cb_tsDefaultTimeout.GetValue()
        rerunFailedTestCaseOnly = self.cb_tsRerun.GetValue()
        if suiteName != '':
            suiteInfo['suiteName'] = suiteName
        else:
            wx.MessageBox(u"测试套件名称不能为空!")
            return
        if description != '':
            suiteInfo['description'] = description
        else:
            wx.MessageBox(u"套件描述信息不能为空!")
            return
        if pageLoadTimeout != '':
            suiteInfo['pageLoadTimeout'] = str(pageLoadTimeout)
        else:
            wx.MessageBox(u"超时时间不能为空!")
            return
        suiteInfo['numberOfRerun']=str(numberOfRerun)
        if pageLoadTimeoutDefault is True:
            suiteInfo['pageLoadTimeoutDefault']="true"
        else:
            suiteInfo['pageLoadTimeoutDefault']="false"
        if rerunFailedTestCaseOnly is True:
            suiteInfo['rerunFailedTestCasesOnly']="true"
        else:
            suiteInfo['rerunFailedTestCasesOnly']="false"

        try:
            commonUtils.create_suite_xml( suiteInfo, self.caseList )
        except Exception as e:
            wx.MessageBox(u'创建TestSuite失败![%s]'%(e.message))
            logging.error(e.message)
            return
        logging.info(u"创建TestSuite[%s]成功!"%(self.tc_tsName.GetValue()))
        wx.MessageBox(u"TestSuite创建成功!")

class AutoTestPanel( wx.Panel ):
    def __init__( self, parent, id=wx.ID_ANY, name="autoSet" ):
        wx.Panel.__init__( self, parent=parent, id=id, name=name )
        self.scheduler = None
        self.jobA = None
        self.jobB = None
        self.run_flag = False
        self.threads=[]
        self.process=[]
        self.il=wx.ImageList( 16, 16 )
        if os.path.exists( "images/yes.png" ):
            bmp = wx.Bitmap( "images/yes.png", wx.BITMAP_TYPE_PNG )
            self.il.Add( bmp )
        if os.path.exists( "images/no.png" ):
            bmp = wx.Bitmap( "images/no.png", wx.BITMAP_TYPE_PNG )
            self.il.Add( bmp )
        if os.path.exists( "images/check.png"):
            bmp = wx.Bitmap( "images/check.png", wx.BITMAP_TYPE_PNG )
            self.il.Add(bmp)

        box = wx.BoxSizer( wx.VERTICAL )

        choices=[(u'模拟测试环境','VIRT'),(u'敏捷测试环境','AGILE'), (u'单元测试环境','DEVP'),(u'集成测试SIT1','SIT1'),(u'集成测试SIT2','SIT2'), (u'验收测试UAT1','UAT1'),(u'验收测试UAT2','UAT2')]

        lab_Env = wx.StaticText( self, -1, label=u'选择测试环境' )
        self.ctrl_Env = ctrl_Env = wx.ComboBox( self, id=wx.ID_ANY, choices=[] )
        #lab_PackAddr = wx.StaticText( self, -1, label=u'测试拉包地址' )
        #self.ctrl_PackAddr = ctrl_PackAddr = wx.TextCtrl( self, -1, value="http://192.168.1.222:8080/" )
        lab_order = wx.StaticText( self, -1, label=u'队列中') 
        self.ctrl_order = ctrl_order = wx.ListCtrl( self, id=wx.ID_ANY, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES )
        lab_over = wx.StaticText( self, -1, label=u'进行中') 
        self.ctrl_over = ctrl_over = wx.ListCtrl( self, id=wx.ID_ANY, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES )
        self.log = log = wx.ListCtrl( self, id=wx.ID_ANY, style=wx.LC_REPORT )
        lab_slider = wx.StaticText( self, id=wx.ID_ANY, label=u'案例覆盖率' )
        self.slider = slider = wx.Slider( self, id=wx.ID_ANY, value=0, minValue=0, maxValue=100, style=wx.SL_HORIZONTAL|wx.SL_AUTOTICKS|wx.SL_LABELS )

        for item in choices:
            self.ctrl_Env.Append( item[0], item[1] )
        self.ctrl_Env.SetSelection(0)

        ctrl_order.InsertColumn(0, u"测试包名")
        ctrl_order.InsertColumn(1, u"版本")
        ctrl_order.InsertColumn(2, u"状态")

        ctrl_over.InsertColumn(0, u"测试包名")
        ctrl_over.InsertColumn(1, u"完成时间")
        ctrl_over.InsertColumn(2, u"完成次数")

        self.log.AssignImageList( self.il, wx.IMAGE_LIST_SMALL )
        log.InsertColumn( 0, u'校验项', width=360, format=wx.LIST_MASK_TEXT^wx.LIST_MASK_IMAGE )
        log.InsertColumn( 1, u'状态', format=wx.LIST_MASK_IMAGE^wx.LIST_MASK_TEXT )

        self.btn_check = btn_check = wx.Button( self, id=wx.ID_ANY, label=u'环境检测' )
        self.btn_start = btn_start = wx.Button( self, id=wx.ID_ANY, label=u'自动任务启动' )
        self.btn_stop = btn_stop = wx.Button( self, id=wx.ID_ANY, label=u'自动任务停止' )

        if os.path.exists("images/checkAuto.png"):
            bmp = wx.Bitmap( "images/checkAuto.png", wx.BITMAP_TYPE_PNG )
            self.btn_check.SetBitmap( bmp )
        if os.path.exists("images/startAuto.png"):
            bmp = wx.Bitmap( "images/startAuto.png", wx.BITMAP_TYPE_PNG )
            self.btn_start.SetBitmap( bmp )
        if os.path.exists("images/stopAuto.png"):
            bmp = wx.Bitmap( "images/stopAuto.png", wx.BITMAP_TYPE_PNG )
            self.btn_stop.SetBitmap( bmp )

        btn_stop.Enable(False)
        btn_start.Enable(False)

        flex = wx.FlexGridSizer( 2, 2, 5, 5 )
        flex.AddMany([
            (lab_Env, 0, wx.ALL|wx.ALIGN_CENTER, 3),
            (ctrl_Env, 1, wx.ALL|wx.EXPAND, 3),
            (lab_slider, 0, wx.ALL|wx.ALIGN_CENTER, 3 ),
            (slider, 0, wx.ALL|wx.EXPAND, 3 )
            #(lab_PackAddr, 0, wx.ALL|wx.ALIGN_CENTER, 3 ),
            #(ctrl_PackAddr, 1, wx.ALL|wx.EXPAND, 3 )
            ])

        flex.AddGrowableCol( 1, 1 )

        vbox_left = wx.BoxSizer( wx.VERTICAL )
        vbox_left.AddMany([
            (lab_order, 0, wx.ALL|wx.ALIGN_CENTER, 3),
            (ctrl_order, 1, wx.EXPAND|wx.ALL, 3 )
            ])
        vbox_right = wx.BoxSizer( wx.VERTICAL )
        vbox_right.AddMany([
            (lab_over, 0, wx.ALL|wx.ALIGN_CENTER, 3),
            (ctrl_over, 1, wx.EXPAND|wx.ALL, 3 )
            ])

        hbox = wx.BoxSizer( wx.HORIZONTAL )
        hbox.Add( vbox_left, 1, wx.EXPAND|wx.ALL, 3 )
        hbox.Add( vbox_right, 1, wx.EXPAND|wx.ALL, 3 )
        vbox_m = wx.BoxSizer( wx.VERTICAL )
        vbox_m.Add( hbox, 1, wx.EXPAND|wx.ALL, 3 )
        vbox_m.Add( log, 1, wx.EXPAND|wx.ALL, 9 )


        hfunc = wx.BoxSizer( wx.HORIZONTAL )
        hfunc.AddMany([
            (btn_check, 0, wx.ALIGN_CENTER|wx.ALL, 5),
            (btn_start, 0, wx.ALIGN_CENTER|wx.ALL, 5),
            (btn_stop, 0, wx.ALIGN_CENTER|wx.ALL, 5)
            ])

        box.Add( flex, 0, wx.EXPAND, 3 )
        box.Add( vbox_m, 1, wx.EXPAND|wx.ALL, 3 )
        box.Add( hfunc, 0, wx.ALIGN_CENTER|wx.ALL, 3 )

        self.SetSizer( box )

        wx.CallAfter( self.InitPack )

        self.Bind( wx.EVT_BUTTON, self.OnStart, btn_start )
        self.Bind( wx.EVT_BUTTON, self.OnStop, btn_stop )
        self.Bind( wx.EVT_BUTTON, self.OnCheck, btn_check )
        self.Bind( wx.EVT_SCROLL_CHANGED, self.OnSlider, slider )
        self.Bind( wx.EVT_COMBOBOX, self.OnChangeEnv, self.ctrl_Env )
    def writeRecord( self, message='', imageId=1, flag=True ):
        #flag True 新建条目 False 更新条目
        #新增 writeRecord( message='xxx')
        idx = self.log.GetItemCount()
        if flag is True:
            #创建一个条目
            self.log.InsertItem( idx, message, imageIndex=2 )
            self.log.SetItem( idx, 1, '', imageId = imageId )
        else:
            #更新上一个条目
            if idx >= 1:
                self.log.SetItem( idx-1, 1, message, imageId = imageId )
        self.log.Update()
    def InitPack( self ):
        AutoTask.InitPackStatus()
        env = self.ctrl_Env.GetClientData( self.ctrl_Env.GetSelection() )
        self.ShowQueuePack( pack_env = env )
    def OnChangeEnv( self, event ):
        env = self.ctrl_Env.GetClientData( self.ctrl_Env.GetSelection() )
        print env
        self.ShowQueuePack( pack_env = env )
    def GetSlider( self ):
        return self.slider
    def SetSliderMax( count ):
        self.slider.SetMaxValue(count)
        self.slider.Update()

    def ShowQueuePack( self, pack_env ):
        self.ctrl_order.DeleteAllItems()
        #显示自动化测试等待队列
        result = AutoTask.GetQueuePack( pack_env=pack_env )
        i = 0
        for item in result:
            self.ctrl_order.InsertItem( i, str(item.get('pack_name')) ) 
            self.ctrl_order.SetItem( i, 1, str(item.get('pack_version')) ) 
            self.ctrl_order.SetItem( i, 2,  u'队列中') 
            i=i+1
    def ShowRunningPack( self, pack_env ):
        self.ctrl_over.DeleteAllItems()
        isRunning, result = AutoTask.hasRunning( pack_env=pack_env )
        if isRunning is True:
            #有正在运行的自动化测试项目
            try:
                j = 0
                for item in result:
                    self.ctrl_over.InsertItem( j, str(item.get('pack_name')) )
                    self.ctrl_over.SetItem( j, 1, str(item.get('pack_version')) )
                    self.ctrl_over.SetItem( j, 2, u'进行中' )
                    j=j+1
            except Exception as e:
                logging.error(e.message)
    def ShowPack( self, pack_env ):
        try:
            isRunning,result = AutoTask.hasRunning( pack_env=pack_env )
            if isRunning is False:
                #无正在运行的自动化测试项目
                try:
                    pack_info = AutoTask.GetQueuePack( pack_env = pack_env )
                    tree = wx.GetApp().GetTopWindow().FindWindowByName("tree")
                    device = AutoTask.GetDevice()
                    maxCase = tree.GetCaseCount()
                    deviceId = device.get('deviceId')
                    if len(pack_info) > 0 and deviceId is not None and maxCase > 0:
                        thread = RunningThread(packInfo=pack_info[0], maxCase=maxCase, deviceId=deviceId, log=self.log, window=self)
                        thread.start()
                        self.threads.append( thread )
                except Exception as e:
                    for thread in self.threads:
                        thread.terminate()
                    logging.error(e.message)
        except Exception as e:
            logging.log(e.message)

    def UpdatePack( self, url, env ):
        try:
            #显示队列
            self.ShowQueuePack( pack_env = env )

            #更新运行包
            self.ShowRunningPack( pack_env = env )

            #更新流程
            AutoTask.updatePack( url=url, env=env)

            self.ShowPack( pack_env=env )

        except Exception as e:
            logging.warning(e.message)

    def ClearPack( self ):
        try:
            AutoTask.clearPack()
        except Exception as e:
            logging.warn(e.message)

    def OnSlider( self, event, count=None ):
        tree = self.GetParent().FindWindowByName("tree")
        if tree:
            self.slider.SetMax(tree.GetCaseCount())
        if count is not None:
            self.slider.SetValue(count)
        self.AddCase( int(self.slider.GetValue()) )

    def AddCase( self, count ):
        #添加测试案例
        tree = self.GetParent().FindWindowByName("tree")
        data=[]
        rootId = tree.GetRootItem()
        items = tree.GetItemData( rootId )
        runData={"level3":{"data":[], "count":0}, "level2":{"data":[],"count":0}, "level1":{"data":[],"count":0}, "level0":{"data":[],"count":0}}
        dataCount=0
        for key in items:
            for item in items.get(key):
                if item.get('weight') == 3:
                    runData.get('level3').get('data').append( item )
                    runData.get('level3')['count'] = runData.get('level3').get('count') + 1
                elif item.get('weight') == 2:
                    runData.get('level2').get('data').append( item )
                    runData.get('level2')['count'] = runData.get('level2').get('count') + 1
                elif item.get('weight') == 1:
                    runData.get('level1').get('data').append( item )
                    runData.get('level1')['count'] = runData.get('level1').get('count') + 1
                elif item.get('weight') == 0:
                    runData.get('level0').get('data').append( item )
                    runData.get('level0')['count'] = runData.get('level0').get('count') + 1
        data = runData.get('level3').get('data')
        dataCount = runData.get('level3').get('count')
        if dataCount < count:
            level2 = int((count-dataCount)*0.8)
            level1 = int((count-dataCount)*0.2)
            if level2 <= runData.get('level2').get('count'):
                data.extend(random.sample( runData.get('level2').get('data'), level2 ))
            else:
                data.extend(runData.get('level2').get('data'))
            if level1 <= runData.get('level1').get('count'):
                data.extend(random.sample( runData.get('level1').get('data'), level1 ))
            else:
                data.extend(runData.get('level1').get('data'))

        if self.GetParent().FindWindowByName("list"):
            self.GetParent().FindWindowByName("list").InitData( data, auto=True )
            #self.GetParent().FindWindowByName("list").Refresh()

    def UpdateData( self ):
        #更新测试数据
        lc = self.GetParent().FindWindowByName("list")
        if lc is not None:
            lc.OnTemp( wx.CommandEvent( wx.EVT_MENU.typeId ), auto=True )

    def CreateTestSuite( self, packName=None, packDesc=None ):
        #创建测试套件
        suiteInfo = {}
        caseList=[]
        lc = self.GetParent().FindWindowByName("list")
        sel = -1
        while True:
            sel = lc.GetNextItem( sel, wx.LIST_NEXT_ALL )
            if sel == -1:
                break
            caseList.append(lc.GetItemText( sel, 6 ))
        suiteName = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if packName is not None:
            suiteName = packName+"_"+suiteName

        description = 'TestSuite:%s'%(suiteName)
        if packDesc is not None:
            description=packDesc

        numberOfRerun = 0
        pageLoadTimeout = 30
        pageLoadTimeoutDefault = "true"
        rerunFailedTestCasesOnly = "false" 

        suiteInfo['suiteName'] = suiteName
        suiteInfo['description'] = description
        suiteInfo['pageLoadTimeout'] = str(pageLoadTimeout)
        suiteInfo['numberOfRerun']=str(numberOfRerun)
        suiteInfo['pageLoadTimeoutDefault']=pageLoadTimeoutDefault
        suiteInfo['rerunFailedTestCasesOnly']=rerunFailedTestCasesOnly

        ret = commonUtils.create_suite_xml( suiteInfo, caseList )
        if ret[0] != 0:
            raise Exception("创建测试套件失败[%s]"%(e.message))
        return suiteName

    def SyncDevice( self, project_path, deviceId ):
        #同步运行设备到项目运行环境
        kv={"deviceId":deviceId}
        result = commonUtils.SetDeviceInfo(project_path, kv )
        if result[0] != 0:
            raise Exception(result[1])

    def autoRun( self, project_path, suiteName, deviceId ):
        #自动运行TestSuite
        katalon_exe = commonUtils.ConfigRead(CONFIG_FILE, "MBKAUTOTEST", "katalon_exe" )
        projectFile = os.path.join( project_path, "MBKAutoTest.prj")
        suitePath = os.path.join( "Test Suites", suiteName ).replace("\\","/")
        cmd = '%s -runMode=console -consoleLog  -projectPath="%s" -statusDelay=95000 -retry=0 -testSuitePath="%s" -deviceId="%s" -browserType="Android"'%(katalon_exe, projectFile, suitePath, deviceId)
        #p = subprocess.Popen( cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True )
        #self.process.append(p)
        try:
            output = subprocess.check_output( cmd, shell=True)
            print output
        except subprocess.CalledProcessError as e:
            raise Exception(e.message)

    def Running( self, suiteName, deviceId ):
        #运行测试套件
        projectPath = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
        suitePath = os.path.join( "Test Suites", suiteName ).replace("\\","/")
        cmd = 'katalon -runMode=console -consoleLog  -projectPath="%s" -statusDelay=95000 -retry=0 -testSuitePath="%s" -deviceId="%s" -browserType="Android"'%(projectFile, suitePath, deviceId)
        p = subprocess.Popen( cmd, stdout=subprocess.PIPE, stderr = subprocess.STDOUT, shell=True )
        output,outerr = p.communicate()
        print output,outerr
        if p.returncode != 0:
            raise Exception(u"执行失败")

    def UpdatePackStatus( self, packInfo, status=2 ):
        #更新测试包状态
        if packInfo is not None:
            packName = packInfo.get('pack_name')
            packVersion = packInfo.get('pack_version')
            packEnv = packInfo.get('pack_env')
            packType = packInfo.get('pack_type')

            try:
                db = web.database( dbn="sqlite", db=AutoTask.dbname )
                db.update( "packrun", where="pack_name=$pack_name and pack_version=$pack_version and pack_type=$pack_type and pack_env=$pack_env", pack_state=status, vars={'pack_name':packName,\
                    'pack_version':packVersion, 'pack_type':packType, 'pack_env':packEnv})
            except Exception as e:
                raise Exception(e.message)

    def checkProcess( self, env ):
        for process in self.process:
            if process.poll() is None:
                return
            else:
                #找到正在运行的测试包信息
                if process in self.process:
                    self.process.remove(process)
                (flag, result) = AutoTask.hasRunning( pack_env=env )
                message=""
                resultFlag = True
                updateFlag = True
                if flag is True:
                    packInfo = result[0]
                    if process.poll() == 0:
                        try:
                            #更新测试包状态
                            self.UpdatePackStatus( packInfo, status='2' )
                        except Exception as e:
                            updateFlag = False
                            raise Exception(e.message)
                    else:
                        resultFlag = False
                        try:
                            #更新测试包状态
                            self.UpdatePackStatus( packInfo, status='0' )
                        except Exception as e:
                            updateFlag = False
                            raise Exception(e.message)
                if resultFlag is True:
                    self.writeRecord( u'测试结果', imageId=0 )
                else:
                    self.writeRecord( u'测试结果', imageId=1 )
                if updateFlag is True:
                    self.writeRecord( u'更新测试包状态', imageId=0 )
                else:
                    self.writeRecord( u'更新测试包状态', imageId=1 )

    def OnCheck( self, event ):
        work = CheckEnvironThread( self.log, self )
        work.start()

    def OnStart( self, event ):
        env = self.ctrl_Env.GetClientData(self.ctrl_Env.GetSelection())
        url = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "pack_url" )
        projectPath = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )

        trigger_s5 = IntervalTrigger(seconds=5)
        trigger_s10 = IntervalTrigger(seconds=10)
        trigger_s20 = IntervalTrigger(seconds=20)
        trigger_s300 = IntervalTrigger(seconds=300)

        #更新包
        self.scheduler = BackgroundScheduler()
        self.jobA = self.scheduler.add_job( self.UpdatePack, trigger_s10, kwargs={"url":url, "env":env} )
        self.jobB = self.scheduler.add_job( self.checkProcess, trigger_s5, kwargs={"env":env} )
        if self.scheduler.running is False:
            self.scheduler.start()
        else:
            self.scheduler.resume()

        print self.scheduler.get_jobs()

        self.btn_start.Enable(False)
        self.ctrl_Env.Enable(False)
        self.btn_stop.Enable(True)
        self.run_flag=True
    def OnStop( self, event ):
        try:
            env = self.ctrl_Env.GetClientData( self.ctrl_Env.GetSelection() )
            AutoTask.InitPackStatus()
            self.ShowQueuePack( pack_env = env )
            self.ShowRunningPack( pack_env = env )
            self.scheduler.shutdown(wait=False)
            self.scheduler.remove_all_jobs()
        except Exception as e:
            logging.error(u'初始测试包状态失败')
        finally:
            if self.scheduler.running:
                self.scheduler.shutdown(wait=False)
                self.scheduler.remove_all_jobs()
                self.scheduler.pause()
        self.btn_start.Enable(False)
        self.btn_stop.Enable(False)
        self.ctrl_Env.Enable(True)
        self.run_flag=False

if __name__ == '__main__':
    app = wx.App()
    frame = AutoFrame( parent=None, style=wx.MAXIMIZE|wx.DEFAULT_FRAME_STYLE )
    frame.CenterOnParent()
    app.SetTopWindow(frame)
    frame.ShowFullScreen(True, wx.FULLSCREEN_NOSTATUSBAR)
    app.MainLoop()
