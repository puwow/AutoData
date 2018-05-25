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

        #У����������
        message=''
        imageIndex=0
        try:
            AutoTask.checkServer( timeout=5 )
        except Exception as e:
            flag=False
            imageIndex=1
        finally:
            self.window.writeRecord( u'У��汾������', imageId=imageIndex)

        #У��SVN
        message=''
        imageIndex=0
        svn_url = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "svn_url" )
        try:
            AutoTask.checkSVN()
        except Exception as e:
            flag=False
            imageIndex=1
        finally:
            self.window.writeRecord( u'У��SVN����', imageId=imageIndex )

        #У��Katalon
        katalonFlag = True
        message=''
        imageIndex=0
        katalonFlag = AutoTask.checkKatalon()
        if katalonFlag is False:
            flag=False
            imageIndex=1

        self.window.writeRecord( u'У��Katalon����', imageId=imageIndex )

        #У���ֻ��豸
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
            self.window.writeRecord(u'У���ֻ��豸����', imageId=imageIndex )

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
        #�Զ���������Ŀ·��
        project_path = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
        #������������
        minCase=1
        self.maxCase=1
        count=0
        if self.maxCase <= minCase:
            count = random.randint( self.maxCase, minCase )
        else:
            count = random.randint( minCase, self.maxCase )

        #����Slider�ؼ��¼���������count������
        wx.PostEvent( self.window, wx.CommandEvent(wx.EVT_SCROLL_CHANGED.typeId, self.window.slider.GetId()) )

        if self.packInfo is None:
            self.window.writeRecord( u'û����Ҫ���Եĳ����' )
            return

        packName = self.packInfo.get('pack_name')
        packDesc = self.packInfo.get('pack_desc')

        idx = 0
        try:
            #���°�״̬
            self.window.UpdatePackStatus( self.packInfo, status='1' )
            #ˢ�°�����
            self.window.ShowQueuePack( pack_env = self.window.ctrl_Env.GetClientData( self.window.ctrl_Env.GetSelection() ) )
            #��ʾ���ڲ����еİ�
            self.window.ShowRunningPack( pack_env = self.window.ctrl_Env.GetClientData( self.window.ctrl_Env.GetSelection() ) )
            #���°�������
            self.window.UpdateData()

            self.window.writeRecord( u'���°�������', imageId=0 )
            try:
                #���ɲ����׼�
                suiteName = self.window.CreateTestSuite( packName=packName, packDesc = packDesc )
                self.window.writeRecord( u'���ɲ����׼�', imageId=0 )
                try:
                    #ͬ�������豸
                    self.window.SyncDevice( project_path = project_path, deviceId=self.deviceId )
                    self.window.writeRecord( u'ͬ�������豸', imageId=0 )
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
        #������Ϣ
        self.data = data
        #��������
        self.data2 = data2
        #�������������Ƿ����
        self.notEqual = False
        self.Bind( wx.EVT_CONTEXT_MENU, self.OnRightClick )
        self.selected = -1 #Ĭ��������
        #��ѡ�����
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
                #��������
                i=0
                for var in variables:
                    #�����ֵ�
                    if var[1] is not None:
                        self.InsertColumn( i, var[1]+"("+var[0]+")" )
                    else:
                        self.InsertColumn( i, var[0] )
                    i=i+1
                    reg = re.search(r'(\d).+(\d)', var[2])
                    if reg:
                         #������TestCase�а󶨵���
                         n = int(reg.groups()[0])
                         self.selected = n
        self.InitData(n, test_case_info.get('name'))
    def isNotE( self ):
        return self.notEqual
    def InitData( self, n, name ):
        #��ʼ������, n�Ǵ�tc��ȡ�õ�Ĭ�����������������ʵ��������1
        ccol = self.GetColumnCount()
        row = 0
        notEqual = False
        for record in self.data2[1:]:
            if len(record) != ccol:
                #�����TC�л�ȡ�ı�������Excel�е�����������һ�£�������������
                logging.warn(u'���԰���[%s]�����������һ��'%(name))
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
        #�ṩ���ݸ�Excel��
        data=[]
        udata=[]
        for col in range( self.GetColumnCount() ):
            column = self.GetColumn( col )
            udata.append(column.GetText())
        data.append(udata)
        sel = -1
        while True:
            #�б���
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
        addMenu = wx.MenuItem( menu, id=self.addMenuId, text=u"�������", helpString=u'���һ������' )
        deleteMenu = wx.MenuItem( menu, id=self.deleteMenuId, text=u"ɾ������", helpString=u'ɾ��һ������')
        defaultMenu = wx.MenuItem( menu, id=self.defaultMenuId, text=u'��ΪĬ��ֵ', helpString=u'���ø�����¼ΪĬ��ֵ')
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
            logging.info(u"ɾ����¼[%s]"%(self.GetItemText( sel, 1 )))
            self.DeleteItem( sel )
    def OnSetDefault( self, event ):
        sel = -1
        while True:
            sel = self.GetNextItem( sel, wx.LIST_NEXT_ALL )
            if sel == -1:
                break
            self.SetItemTextColour(sel, self.ncolour )
        row = self.GetFirstSelected()
        #����Ĭ��ֵ��ListCtrl�е�ֵ��0�𲽣�Excel���е�ֵ��1��
        if row <=0:
            self.selected = 1
        if row >=0:
            self.selected = row + 1
            self.SetItemTextColour(row, self.gcolour)
        logging.info(u'���õ�[%d]����¼ΪĬ��ִ������!'%(self.selected))
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
            root = self.AddRoot(u"��Ŀ��������", image=4, selImage=4, data=dict_cases)
            if dict_cases:
                count=0
                for key in dict_cases.keys():
                    logging.info(key)
                    desc=''
                    key_val=key
                    if she is not None:
                        #�洢Key��ȡ��Keyֵ
                        if key not in she.keys():
                            she[str(key)]={'desc':'', 'childs':[]}
                        else:
                            desc=she[str(key)]['desc']
                    #�����ڵ�
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
                        #�����ڵ�
                        index = 1
                        if case.get("desc") == u"δ���":
                            index=3
                        self.AppendItem( node, case.get("name")+"("+case.get("desc")+")", image=index, selImage=2, data=[case] )
                        count=count+1
                logging.info(u'�����ز��԰���[%d]��!'%(count))
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
                logging.info(u'������ADD����[%s]'%(self.GetItemText(item)))
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
        addMenu = wx.MenuItem( menu, id=self.addMenuId, text=u'���Case', helpString=u'���һ����¼' )
        freshMenu = wx.MenuItem( menu, id=self.freshMenuId, text=u'ˢ��Case', helpString=u'ˢ�¼�¼' )
        expandMenu = wx.MenuItem( menu, id=self.expandMenuId, text=u'ȫ��չ��', helpString=u'չ��ȫ������' )
        collapseMenu = wx.MenuItem( menu, id=self.collapseMenuId, text=u'ȫ���۵�', helpString=u'�۵�ȫ������' )
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
        self.InsertColumn( 0, u"�������" )
        self.InsertColumn( 1, u"��������", width=160 )
        self.InsertColumn( 2, u"��������", width=240 )
        self.InsertColumn( 3, u"��θ���", width=80 )
        self.InsertColumn( 4, u"������", width=80)
        self.InsertColumn( 5, u"Ȩ��", width=80)
        self.InsertColumn( 6, u"�ű�·��", width=-1)


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
        busy = wx.BusyInfo(u"���ݼ��أ����Ժ�...", parent=self)
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
            #�ֶ����
            for item in data:
                if item not in self.userAdd:
                    self.userAdd.append( item )
        else:
            #�Զ����
            for item in data:
                if item  not in self.autoAdd:
                    if item not in allAdd:
                        self.autoAdd.append( item )
        allAdd = allAdd + self.autoAdd

        i=self.GetItemCount()
        for item in  sorted( allAdd, cmp=lambda x,y:cmp(x.get('weight'),y.get('weight')), reverse=True ):
            #Ȩ��ת��
            weight_info = ''
            weight = item.get('weight')
            if weight == 0:
                weight_info = u'����Ҫ'
            elif weight == 1:
                weight_info = u'һ��'
            elif weight == 2:
                weight_info = u'�е�'
            elif weight == 3:
                weight_info = u'��Ҫ'
            self.InsertItem( i, str(i+1) )
            self.SetItem( i, 1, item.get("name") )
            self.SetItem( i, 2, item.get("desc") )

            if item.get("desc") == u"δ���":
                self.SetItemBackgroundColour(i, "yellow")

            #��ȡ��ǰSheetҳ������
            rdata = commonUtils.read_excel( item.get("name") )

            #������
            data_count=0
            if len(rdata) == 1 or len(rdata) == 0:
                data_count=0
            else:
                data_count=len(rdata)-1



            #��θ���
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
            logging.info(u"������ADD����[%s(%s)]"%(item.get("name"),item.get("desc")))
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
        deleteMenu = wx.MenuItem( menu, id=self.deleteMenuId, text=u'ɾ��Case', helpString=u'ɾ��һ����¼' )
        clearMenu = wx.MenuItem( menu, id=self.clearMenuId, text=u'��ռ�¼', helpString=u'��ռ�¼' )
        tempMenu = wx.MenuItem( menu, id=self.tempMenuId, text=u'��������ģ��', helpString=u'���ɱ�ѡ�����Ĳ�������ģ��' )
        suiteMenu = wx.MenuItem( menu, id=self.suiteMenuId, text=u'����TestSuite', helpString=u'�����б�����TestSuite' )
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
            logging.info(u'������DELETE����[%s(%s)]'%(case_name, self.GetItemText(sel, 2)))
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
            wx.MessageBox(u"���ݸ������")
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
            wx.MessageBox( u'������Ӳ��԰���' )
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
        busy = wx.BusyInfo(u"���ڶ�ȡ�����ļ�����Ⱥ�...", parent=parent)
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
    def __init__( self, parent=None, title=u'�Ի���', size=(600,480), style=wx.DEFAULT_DIALOG_STYLE ):
        wx.Dialog.__init__( self, parent=parent, title=title, size=size, style=style )

        panel = wx.Panel( self, -1 )
        self.pg = wxpg.PropertyGridManager( panel, style=wxpg.PG_SPLITTER_AUTO_CENTER|wxpg.PG_AUTO_SORT|wxpg.PG_TOOLBAR )
        self.pg.SetExtraStyle( wxpg.PG_EX_HELP_AS_TOOLTIPS )
        self.pg.AddPage(u"�ֻ������Զ�������ȫ������")
        self.pg.Append( wxpg.DirProperty( u"��ĿĿ¼", label="project_path", value="" ) )
        self.pg.Append( wxpg.DirProperty( u"���ݴ洢·��", label="data_path", value="" ) )
        self.pg.Append( wxpg.BoolProperty( u'���ݵ��ļ�', name="single_excel", value=True ) )
        self.pg.Append( wxpg.StringProperty( u'�����ļ���', name="data_name", value=u"TestData.xls" ) )
        self.pg.Append( wxpg.FileProperty( u'Katalon����', name="katalon_exe", value="" ) )
        self.pg.Append( wxpg.StringProperty( u'��ĿSNV��ַ', name="svn_url", value="svn://14.16.18.5/MBPSYS/trunk/MBKAutoTest" ) )
        self.pg.Append( wxpg.StringProperty( u'���԰������ַ', name="pack_url", value="http://13.239.21.170:8080/" ) )
        self.pg.Append( wxpg.FileProperty( u'adb����', name="adb_exe", value="" ) )

        vbox = wx.BoxSizer( wx.VERTICAL )
        vbox.Add( self.pg, 1, wx.EXPAND|wx.ALL, 3 )

        btn_save = wx.Button( panel, -1, u'����' )
        btn_close = wx.Button( panel, -1, u'�ر�' )
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
            logging.info(u'����ȫ������[%s:%s]'%(key, d[key]))
        wx.MessageBox(u"���óɹ�")
        if self.GetParent():
            if self.GetParent().FindWindowByName("list"):
                self.GetParent().FindWindowByName("list").InitData([])
            if self.GetParent().FindWindowByName("tree"):
                self.GetParent().FindWindowByName("tree").InitData()
        self.Destroy()
    def OnClose( self, event ):
        self.Destroy()

class AutoFrame( wx.Frame ):
    def __init__( self, parent=None, title=u"�ֻ������Զ����������ݹ���", size=(800,600), style=wx.DEFAULT_FRAME_STYLE ):
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
                Caption(u'���԰���(Test Case)').
                CloseButton(False).
                PaneBorder(False).
                Name("TestCase"))
        self.mgr.AddPane( self.autoPanel, wx.aui.AuiPaneInfo().Right().
                BestSize(600,-1).
                MinSize(600, -1).
                Floatable(False).
                Caption(u'�Զ�������').
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
                Caption(u'�����׼�(Test Suite)').
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
                Caption(u'��������(Test Data)').
                PaneBorder(False).
                CloseButton(False).
                Name('TestData'))
        self.mgr.Update()

        self.CreateMenuBar()
        self.Bind( wx.EVT_CLOSE, self.OnClose )
        #���������
        if os.path.exists("MBKAuto.exe"):
            self.SetIcon(wx.Icon("MBKAuto.exe", wx.BITMAP_TYPE_ICO) )
        logging.info(u"Ӧ���������...")
    def OnClose( self, event ):
        logging.info(u"Ӧ�ùر�...")
        if she:
            she.close()
        self.mgr.UnInit()
        self.Destroy()
    def CreateMenuBar( self ):
        menuBar = wx.MenuBar()

        menuGlobal = wx.Menu()
        menuSetting = wx.MenuItem( menuGlobal, id=wx.ID_ANY, text=u'����(&S)', helpString=u'����ȫ�ֱ���' )
        if os.path.exists( "images/setting24.png" ):
            bmp = wx.Bitmap( "images/setting24.png", wx.BITMAP_TYPE_PNG )
            menuSetting.SetBitmap( bmp )
        menuGlobal.Append( menuSetting )

        '''
        menuAuto = wx.MenuItem( menuGlobal, id=wx.ID_ANY, text=u'�Զ�������(&T)', helpString=u'�Զ�������')
        if os.path.exists( "images/auto24.png"):
            bmp = wx.Bitmap( "images/auto24.png", wx.BITMAP_TYPE_PNG)
            menuAuto.SetBitmap( bmp )
        menuGlobal.Append(menuAuto)
        '''

        menuHelp = wx.Menu()
        menuAbout = wx.MenuItem( menuHelp, id=wx.ID_ANY, text=u'����(&A)', helpString=u'������Ϣ' )
        if os.path.exists( "images/about24.png" ):
            bmp = wx.Bitmap( "images/about24.png", wx.BITMAP_TYPE_PNG )
            menuAbout.SetBitmap( bmp )
        menuHelp.Append( menuAbout )

        menuMemo = wx.MenuItem(menuHelp, id=wx.ID_ANY, text=u'��ע(&M)', helpString=u'��ע��Ϣ' )
        if os.path.exists("images/memo24.png"):
            bmp = wx.Bitmap( "images/memo24.png")
            menuMemo.SetBitmap( bmp )
        menuHelp.Append( menuMemo )

        menuBar.Append( menuGlobal, u"ȫ��(&G)" )
        menuBar.Append( menuHelp, u'����(&H)' )
        
        self.SetMenuBar( menuBar )
        self.Bind( wx.EVT_MENU, self.OnSetting, menuSetting )
        self.Bind( wx.EVT_MENU, self.OnAbout, menuAbout )
        self.Bind( wx.EVT_MENU, self.OnMemo, menuMemo )
    def OnSetting( self, event ):
        dialog = SettingDialog( parent=self, title=u'��������', size=(int(self.GetClientSize().width*0.65), int(self.GetClientSize().height*0.65)) )
        dialog.CenterOnParent()
        dialog.ShowModal()
        dialog.Destroy()
    def OnAbout( self, event ):
        info = wx.adv.AboutDialogInfo()
        info.Name=u"�ֻ������Զ����������ݹ�����"
        info.Version=u"1.1.0"
        info.Copyright=u"(c) 2018-2020 �κδ���ͳ����ʹ��Ȩ��"
        info.Description=wordwrap(
                u"��������������Katalon���������Ŀ�ṹ��Ϊ�˱��ڲ������ݵĹ����밸���󶨶����еĶ��ζ��ƿ�������������Ҫ���������������Ե��뷨��Ŀ�����ڽ���������ݵ���Ч����ͼ�����Ϊ��Ԥ���������⡣\n\n���������wxPython GUI�⿪��", 300, wx.ClientDC(self))
        
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
    def __init__( self, parent, id=wx.ID_ANY, title=u'��Ϣ��ע', size=(480,240) ):
        wx.Dialog.__init__( self, parent=parent, id=id, title=title, size=size )
        panel = wx.Panel( self, -1 )

        st_module = wx.StaticText( panel, -1, label=u'����ģ��' )
        st_desc = wx.StaticText( panel, -1, label=u'ģ�鱸ע' )
        st_cases = wx.StaticText( panel, -1, label=u'��������' )
        st_weight = wx.StaticText( panel, -1, label=u'Ȩ��' )
        self.ct_module = ct_module = wx.ComboBox( panel, -1 )
        self.ct_cases = ct_cases = wx.ComboBox( panel, -1 )
        self.ct_desc = ct_desc = wx.TextCtrl( panel, -1 )
        self.ct_weight = ct_weight = wx.Choice( panel, -1, choices=[u'����Ҫ',u'һ��',u'�е�',u'��Ҫ'] )
        save = wx.Button( panel, id=wx.ID_ANY, label=u'����')

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
                    wx.MessageBox(u'����ɹ�!')
                except Exception as e:
                    wx.MessageBox(u'����ʧ��[%s]'%(e.message))
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
    def __init__( self, parent, caseList=None, id=wx.ID_ANY, title=u'TestSuite��װ', size=(540,420) ):
        wx.Dialog.__init__( self, parent=parent, id=id, title=title, size=size )

        self.caseList = caseList
        panel = wx.Panel( self, -1 )

        tz = pytz.timezone('Asia/Shanghai')
        now = datetime.datetime.now(tz)
        sb = wx.StaticBitmap( panel, id=wx.ID_ANY )
        tsName = wx.StaticText( panel, id=wx.ID_ANY, label=u'�����׼�����' )
        self.tc_tsName = wx.TextCtrl( panel, id=wx.ID_ANY, value=now.strftime('%Y%m%d') )
        tsDesc = wx.StaticText( panel, id=wx.ID_ANY, label=u'�����׼�����' )
        self.tc_tsDesc = wx.TextCtrl( panel, id=wx.ID_ANY, style=wx.TE_MULTILINE )
        tsRuntimes = wx.StaticText( panel, id=wx.ID_ANY, label=u'���д���' )
        self.ic_tsRuntimes = wx.lib.intctrl.IntCtrl( panel, value=0, min=0 )
        tsTimeout = wx.StaticText( panel, id=wx.ID_ANY, label=u'��ʱʱ��' )
        self.ic_tsTimeout = wx.lib.intctrl.IntCtrl( panel, id=wx.ID_ANY, value=30, min=30 )
        tsDefaultTimeout = wx.StaticText( panel, id=wx.ID_ANY, label=u'ʹ��Ĭ�ϳ�ʱʱ��')
        self.cb_tsDefaultTimeout = wx.CheckBox( panel, id=wx.ID_ANY )
        tsRerun = wx.StaticText( panel, id=wx.ID_ANY, label=u'���������д�����' )
        self.cb_tsRerun = wx.CheckBox( panel, id=wx.ID_ANY )
        self.cb_tsDefaultTimeout.SetValue(True)

        self.tc_tsDesc.SetValue(now.isoformat())
        btn_run = wx.Button( panel, id=wx.ID_ANY, label=u"����" )
        btn_save = wx.Button( panel, id=wx.ID_ANY, label=u'����' )
        if os.path.exists("images/create32.png"):
            bmp = wx.Bitmap( "images/create32.png", wx.BITMAP_TYPE_PNG )
            btn_save.SetBitmap( bmp )
        if os.path.exists("images/run32.png"):
            bmp = wx.Bitmap( "images/run32.png", wx.BITMAP_TYPE_PNG )
            btn_run.SetBitmap( bmp )

        kata_device_label = wx.StaticText( panel, id=wx.ID_ANY, label=u'KatalonĬ���豸' )
        sys_device_label = wx.StaticText( panel, id=wx.ID_ANY, label=u'��ǰ�����豸' )
        self.kata_device = wx.ComboBox( panel, id=wx.ID_ANY )
        self.sys_device = wx.ComboBox( panel, id=wx.ID_ANY )
        sync = wx.Button( panel, id=wx.ID_ANY, label=u'ͬ��' )

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
            wx.MessageBox(u"δ̽�⵽�ֻ�����")
            return
        if self.sys_device.GetValue() != self.kata_device.GetValue():
            dialog = wx.MessageDialog( self, message=u"�Ƿ񽫵�ǰ�����ֻ���KatalonĬ���ֻ�ͬ��?", caption=u'ͬ����Ϣ', style=wx.OK|wx.CANCEL )
            ret = dialog.ShowModal()
            if ret == wx.ID_OK:
                projectPath = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
                kv={"deviceId":self.sys_device.GetValue()}
                commonUtils.SetDeviceInfo(projectPath, kv )
                self.InitDevice()
    def OnRun( self, event ):
        if self.sys_device.GetValue() != self.kata_device.GetValue():
            wx.MessageBox(u"����ͬ���豸��Ϣ!")
            return
        ret = commonUtils.Executeable("katalon")
        if ret == False:
            wx.MessageBox(u"�뽫����[katalon]������ӵ���������PATH��")
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
            wx.MessageBox(u"ִ�д���!")
            logging.info(u"ִ�д���[%s]!"%(output) )
        logging.info(u"����TestSuite[%s]"%(self.tc_tsName.GetValue()))
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
            wx.MessageBox(u"�����׼����Ʋ���Ϊ��!")
            return
        if description != '':
            suiteInfo['description'] = description
        else:
            wx.MessageBox(u"�׼�������Ϣ����Ϊ��!")
            return
        if pageLoadTimeout != '':
            suiteInfo['pageLoadTimeout'] = str(pageLoadTimeout)
        else:
            wx.MessageBox(u"��ʱʱ�䲻��Ϊ��!")
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
            wx.MessageBox(u'����TestSuiteʧ��![%s]'%(e.message))
            logging.error(e.message)
            return
        logging.info(u"����TestSuite[%s]�ɹ�!"%(self.tc_tsName.GetValue()))
        wx.MessageBox(u"TestSuite�����ɹ�!")

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

        choices=[(u'ģ����Ի���','VIRT'),(u'���ݲ��Ի���','AGILE'), (u'��Ԫ���Ի���','DEVP'),(u'���ɲ���SIT1','SIT1'),(u'���ɲ���SIT2','SIT2'), (u'���ղ���UAT1','UAT1'),(u'���ղ���UAT2','UAT2')]

        lab_Env = wx.StaticText( self, -1, label=u'ѡ����Ի���' )
        self.ctrl_Env = ctrl_Env = wx.ComboBox( self, id=wx.ID_ANY, choices=[] )
        #lab_PackAddr = wx.StaticText( self, -1, label=u'����������ַ' )
        #self.ctrl_PackAddr = ctrl_PackAddr = wx.TextCtrl( self, -1, value="http://192.168.1.222:8080/" )
        lab_order = wx.StaticText( self, -1, label=u'������') 
        self.ctrl_order = ctrl_order = wx.ListCtrl( self, id=wx.ID_ANY, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES )
        lab_over = wx.StaticText( self, -1, label=u'������') 
        self.ctrl_over = ctrl_over = wx.ListCtrl( self, id=wx.ID_ANY, style=wx.LC_REPORT|wx.LC_HRULES|wx.LC_VRULES )
        self.log = log = wx.ListCtrl( self, id=wx.ID_ANY, style=wx.LC_REPORT )
        lab_slider = wx.StaticText( self, id=wx.ID_ANY, label=u'����������' )
        self.slider = slider = wx.Slider( self, id=wx.ID_ANY, value=0, minValue=0, maxValue=100, style=wx.SL_HORIZONTAL|wx.SL_AUTOTICKS|wx.SL_LABELS )

        for item in choices:
            self.ctrl_Env.Append( item[0], item[1] )
        self.ctrl_Env.SetSelection(0)

        ctrl_order.InsertColumn(0, u"���԰���")
        ctrl_order.InsertColumn(1, u"�汾")
        ctrl_order.InsertColumn(2, u"״̬")

        ctrl_over.InsertColumn(0, u"���԰���")
        ctrl_over.InsertColumn(1, u"���ʱ��")
        ctrl_over.InsertColumn(2, u"��ɴ���")

        self.log.AssignImageList( self.il, wx.IMAGE_LIST_SMALL )
        log.InsertColumn( 0, u'У����', width=360, format=wx.LIST_MASK_TEXT^wx.LIST_MASK_IMAGE )
        log.InsertColumn( 1, u'״̬', format=wx.LIST_MASK_IMAGE^wx.LIST_MASK_TEXT )

        self.btn_check = btn_check = wx.Button( self, id=wx.ID_ANY, label=u'�������' )
        self.btn_start = btn_start = wx.Button( self, id=wx.ID_ANY, label=u'�Զ���������' )
        self.btn_stop = btn_stop = wx.Button( self, id=wx.ID_ANY, label=u'�Զ�����ֹͣ' )

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
        #flag True �½���Ŀ False ������Ŀ
        #���� writeRecord( message='xxx')
        idx = self.log.GetItemCount()
        if flag is True:
            #����һ����Ŀ
            self.log.InsertItem( idx, message, imageIndex=2 )
            self.log.SetItem( idx, 1, '', imageId = imageId )
        else:
            #������һ����Ŀ
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
        #��ʾ�Զ������Եȴ�����
        result = AutoTask.GetQueuePack( pack_env=pack_env )
        i = 0
        for item in result:
            self.ctrl_order.InsertItem( i, str(item.get('pack_name')) ) 
            self.ctrl_order.SetItem( i, 1, str(item.get('pack_version')) ) 
            self.ctrl_order.SetItem( i, 2,  u'������') 
            i=i+1
    def ShowRunningPack( self, pack_env ):
        self.ctrl_over.DeleteAllItems()
        isRunning, result = AutoTask.hasRunning( pack_env=pack_env )
        if isRunning is True:
            #���������е��Զ���������Ŀ
            try:
                j = 0
                for item in result:
                    self.ctrl_over.InsertItem( j, str(item.get('pack_name')) )
                    self.ctrl_over.SetItem( j, 1, str(item.get('pack_version')) )
                    self.ctrl_over.SetItem( j, 2, u'������' )
                    j=j+1
            except Exception as e:
                logging.error(e.message)
    def ShowPack( self, pack_env ):
        try:
            isRunning,result = AutoTask.hasRunning( pack_env=pack_env )
            if isRunning is False:
                #���������е��Զ���������Ŀ
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
            #��ʾ����
            self.ShowQueuePack( pack_env = env )

            #�������а�
            self.ShowRunningPack( pack_env = env )

            #��������
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
        #��Ӳ��԰���
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
        #���²�������
        lc = self.GetParent().FindWindowByName("list")
        if lc is not None:
            lc.OnTemp( wx.CommandEvent( wx.EVT_MENU.typeId ), auto=True )

    def CreateTestSuite( self, packName=None, packDesc=None ):
        #���������׼�
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
            raise Exception("���������׼�ʧ��[%s]"%(e.message))
        return suiteName

    def SyncDevice( self, project_path, deviceId ):
        #ͬ�������豸����Ŀ���л���
        kv={"deviceId":deviceId}
        result = commonUtils.SetDeviceInfo(project_path, kv )
        if result[0] != 0:
            raise Exception(result[1])

    def autoRun( self, project_path, suiteName, deviceId ):
        #�Զ�����TestSuite
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
        #���в����׼�
        projectPath = commonUtils.ConfigRead( CONFIG_FILE, "MBKAUTOTEST", "project_path" )
        suitePath = os.path.join( "Test Suites", suiteName ).replace("\\","/")
        cmd = 'katalon -runMode=console -consoleLog  -projectPath="%s" -statusDelay=95000 -retry=0 -testSuitePath="%s" -deviceId="%s" -browserType="Android"'%(projectFile, suitePath, deviceId)
        p = subprocess.Popen( cmd, stdout=subprocess.PIPE, stderr = subprocess.STDOUT, shell=True )
        output,outerr = p.communicate()
        print output,outerr
        if p.returncode != 0:
            raise Exception(u"ִ��ʧ��")

    def UpdatePackStatus( self, packInfo, status=2 ):
        #���²��԰�״̬
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
                #�ҵ��������еĲ��԰���Ϣ
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
                            #���²��԰�״̬
                            self.UpdatePackStatus( packInfo, status='2' )
                        except Exception as e:
                            updateFlag = False
                            raise Exception(e.message)
                    else:
                        resultFlag = False
                        try:
                            #���²��԰�״̬
                            self.UpdatePackStatus( packInfo, status='0' )
                        except Exception as e:
                            updateFlag = False
                            raise Exception(e.message)
                if resultFlag is True:
                    self.writeRecord( u'���Խ��', imageId=0 )
                else:
                    self.writeRecord( u'���Խ��', imageId=1 )
                if updateFlag is True:
                    self.writeRecord( u'���²��԰�״̬', imageId=0 )
                else:
                    self.writeRecord( u'���²��԰�״̬', imageId=1 )

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

        #���°�
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
            logging.error(u'��ʼ���԰�״̬ʧ��')
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
