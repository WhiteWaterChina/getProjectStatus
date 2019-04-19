#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import requests
import re
from bs4 import BeautifulSoup
import xlsxwriter
import os
import time
import types
import datetime
from threading import Thread
import wx
import urllib2
from multiprocessing import Pool
import multiprocessing
import json

num_year = "2019"
ver = "0.5"
address_web = "172.31.2.106:8082"


def get_next(get_data, id_sub, headers_sub, url_sub):
    payload_next_sub = "id={}".format(id_sub)
    get_page = get_data.post(url_sub, headers=headers_sub, data=payload_next_sub)
    data_page = json.loads(get_page.text)
    # print("Get detail info for id:%s with return code %s" % (id, get_page.status_code))
    data_return_dict = {}
    if len(data_page) == 0:
        return None
    else:
        for item_data in data_page:
            id_data = item_data["id"]
            text_data = item_data["text"]
            if "attributes" not in item_data or "testCaseNumber" not in item_data["attributes"]:
                casenumber_data = "None"
            else:
                casenumber_data = item_data["attributes"]["testCaseNumber"]
            data_return_dict["{}".format(id_data)] = {}
            data_return_dict["{}".format(id_data)]["id"] = id_data
            data_return_dict["{}".format(id_data)]["name"] = text_data
            data_return_dict["{}".format(id_data)]["casenumber"] = casenumber_data
            data_return_dict["{}".format(id_data)]["data"] = {}
            data_return_dict["{}".format(id_data)]["parentid"] = id_sub

    return data_return_dict


def add_item_to_dict(value_return, parent_id, dict_sub):
    dict_value_sub = dict_sub
    for key_dict, value_dict in dict_value_sub.items():
        if value_dict["id"] == parent_id:
            value_dict["data"]["{}".format(value_return["id"])] = value_return
            return dict_value_sub
        else:
            if isinstance(value_dict["data"], dict):
                add_item_to_dict(value_return, parent_id, value_dict["data"])
    return dict_value_sub


def add_level():
    pass

def get_detail(get_data, id_testcase_all, flag_status_list):
    print(id_testcase_all)
    get_data_sub = get_data
    url_testcase = "http://{}/iauto_acp/itmsTestCaseN.do/projectConfigTestCaseInfo.view".format(address_web)
    headers_detail = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
        'Content-Length': '4',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Host': '{}'.format(address_web),
        'Origin': 'http://{}'.format(address_web),
        'Referer': 'http://{}/iauto_acp/itmsTestCase.do/testAdmin.view'.format(address_web),
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
        'Upgrade-Insecure-Requests': "1",
    }
    project_id = id_testcase_all.split(":")[1]
    id_testcase = id_testcase_all.split(":")[2]
    querystring_detail = {"projectId": "{}".format(project_id), "configTestCaseId": "{}".format(id_testcase)}
    page_testcase_temp = get_data_sub.get(url_testcase, headers=headers_detail, params=querystring_detail)
    page_testcase = BeautifulSoup(page_testcase_temp.text, "html.parser")
    # bug ID
    bug_id = page_testcase.find("td", text="Bug Id:").parent.find("a").get_text()
    # BUG内容
    bug_content = page_testcase.find("td", text="Bug Content:").parent.find("a").get_text()
    # 备注
    content_bak = page_testcase.find("td", text="备注(content):".decode('gbk')).parent.find("a").get_text()
    produce_temp = re.search(r'var procedureList = (\[\{.*?\}\]);', page_testcase_temp.text).groups()[0]
    produce = json.loads(produce_temp)
    step_list = []
    expect_list = []
    status_list = []
    content_list = []
    if len(produce) != 0:
        for item_step in produce:
            status_step = item_step["result"]
            if status_step in flag_status_list:
                status_list.append(status_step)
                step_list.append(item_step["testProcedure"])
                expect_list.append(item_step["testExpect"])
                content_list.append(item_step["remark"])
    # print(page_testcase)
    # print("zanting")
    get_data_sub.close()
    return id_testcase_all, bug_id, bug_content, content_bak, status_list, step_list, expect_list, content_list


class GetProjectStatus(wx.Frame):
    def __init__(self, parent):

        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"获取ITMS系统上某个项目某个阶段的所有配置的NP和BLOCK的项用例项目-{}".format(ver),
                          pos=wx.DefaultPosition, size=wx.Size(504, 665),
                          style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_APPWORKSPACE))

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel1 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        self.m_panel1.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))

        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_title1 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 1.请输入ITMS系统的用户名和密码！", wx.DefaultPosition,
                                         wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title1.Wrap(-1)

        self.text_title1.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title1.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title1.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer3.Add(self.text_title1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer3, 0, wx.EXPAND, 5)

        gSizer2 = wx.GridSizer(0, 2, 0, 0)

        self.text_username = wx.StaticText(self.m_panel1, wx.ID_ANY, u"用户名", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_username.Wrap(-1)

        self.text_username.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_username.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer2.Add(self.text_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_username = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          0)
        gSizer2.Add(self.input_username, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.text_password = wx.StaticText(self.m_panel1, wx.ID_ANY, u"密码", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_password.Wrap(-1)

        self.text_password.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_password.SetBackgroundColour(wx.Colour(0, 128, 0))

        gSizer2.Add(self.text_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.input_password = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                          wx.TE_PASSWORD)
        gSizer2.Add(self.input_password, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(gSizer2, 0, 0, 5)

        bSizer101 = wx.BoxSizer(wx.VERTICAL)

        self.btn_getprojectname = wx.Button(self.m_panel1, wx.ID_ANY, u"获取所有项目的名称", wx.DefaultPosition, wx.DefaultSize,
                                            0)
        bSizer101.Add(self.btn_getprojectname, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

        bSizer10.Add(bSizer101, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        bSizer19 = wx.BoxSizer(wx.VERTICAL)

        self.text_title11 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 2.请选择要分析的项目！", wx.DefaultPosition,
                                          wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title11.Wrap(-1)

        self.text_title11.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title11.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title11.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer19.Add(self.text_title11, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer4.Add(bSizer19, 0, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer20 = wx.BoxSizer(wx.VERTICAL)

        listbox_projectnameChoices = []
        self.listbox_projectname = wx.ListBox(self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                              listbox_projectnameChoices, 0)
        bSizer20.Add(self.listbox_projectname, 1, wx.ALL | wx.EXPAND, 5)

        bSizer4.Add(bSizer20, 1, wx.EXPAND, 5)

        bSizer16 = wx.BoxSizer(wx.VERTICAL)

        self.btn_getphase = wx.Button(self.m_panel1, wx.ID_ANY, u"获取项目的阶段", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer16.Add(self.btn_getphase, 0, wx.ALL | wx.EXPAND, 5)

        bSizer4.Add(bSizer16, 0, wx.EXPAND, 5)

        bSizer10.Add(bSizer4, 1, wx.EXPAND, 5)

        bSizer12 = wx.BoxSizer(wx.VERTICAL)

        bSizer13 = wx.BoxSizer(wx.VERTICAL)

        self.text_title13 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 3：请选择阶段！", wx.DefaultPosition,
                                          wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title13.Wrap(-1)

        self.text_title13.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title13.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title13.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer13.Add(self.text_title13, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer12.Add(bSizer13, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer15 = wx.BoxSizer(wx.VERTICAL)

        listbox_phaseChoices = []
        self.listbox_phase = wx.ListBox(self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                        listbox_phaseChoices, 0)
        bSizer15.Add(self.listbox_phase, 1, wx.ALL | wx.EXPAND, 5)

        bSizer12.Add(bSizer15, 1, wx.EXPAND, 5)

        bSizer10.Add(bSizer12, 0, wx.EXPAND, 5)

        bSizer21 = wx.BoxSizer(wx.VERTICAL)

        bSizer211 = wx.BoxSizer(wx.VERTICAL)

        self.text_title12 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 4.请点击GO开始导出！或者点击EXIT退出程序！",
                                          wx.DefaultPosition, wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title12.Wrap(-1)

        self.text_title12.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title12.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title12.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer211.Add(self.text_title12, 0, wx.EXPAND, 5)

        bSizer21.Add(bSizer211, 0, wx.EXPAND, 5)

        bSizer22 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self.m_panel1, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer22.Add(self.button_go, 0, wx.ALL, 5)

        self.button_exit = wx.Button(self.m_panel1, wx.ID_ANY, u"EXIT", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer22.Add(self.button_exit, 0, wx.ALL, 5)

        bSizer21.Add(bSizer22, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer21, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        bSizer91 = wx.BoxSizer(wx.VERTICAL)

        self.textctrl_display = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                            wx.DefaultSize, wx.TE_MULTILINE | wx.TE_READONLY)
        bSizer91.Add(self.textctrl_display, 1, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer91, 1, wx.EXPAND, 5)

        self.m_panel1.SetSizer(bSizer10)
        self.m_panel1.Layout()
        bSizer10.Fit(self.m_panel1)
        bSizer2.Add(self.m_panel1, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(bSizer2)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btn_getprojectname.Bind(wx.EVT_BUTTON, self.get_projectname)
        self.btn_getphase.Bind(wx.EVT_BUTTON, self.get_phase)
        self.button_go.Bind(wx.EVT_BUTTON, self.onbutton)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

        self._thread = Thread(target=self.run, args=())
        self._thread.daemon = True

    def close(self, event):
        self.Close()

    def onbutton(self, event):
        self._thread.start()
        self.started = True
        self.button_go = event.GetEventObject()
        self.button_go.Disable()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.textctrl_display.AppendText("完成".decode('gbk') + unicode(t) + "%")
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.textctrl_display.AppendText("%s".decode('gbk') % t)
        self.textctrl_display.AppendText(os.linesep)

    def get_projectname(self, event):
        self.updatedisplay("开始获取项目名称".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))

        # 获取图形界面上输入的信息
        username = self.input_username.GetValue()
        password = self.input_password.GetValue()

        # 登录
        url_login = "http://{}/iauto_acp/login".format(address_web)
        headers_login = {
            'Accept': '*/*',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Content-Length': '32',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Host': '{}'.format(address_web),
            'Origin': 'http://{}'.format(address_web),
            'Referer': 'http://{}/iauto_acp/login.html'.format(address_web),
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
        }
        payload_login = "username={}&password={}".format(username, password)
        get_data = requests.session()
        get_data.post(url_login, data=payload_login, headers=headers_login)

        # 获取所有项目名称
        url_projecttree = "http://{}/iauto_acp/projectAndRound.do/loadProjectTree".format(address_web)
        payload_projecttree = "id=0"
        headers_projecttree = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Content-Length': '4',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Host': '{}'.format(address_web),
            'Origin': 'http://{}'.format(address_web),
            'Referer': 'http://{}/iauto_acp/itmsTestCase.do/testAdmin.view'.format(address_web),
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
        }

        content_projecttree_temp = get_data.post(url_projecttree, data=payload_projecttree,
                                                 headers=headers_projecttree).text
        content_projecttree = json.loads(content_projecttree_temp)
        # project_name_show = []
        if len(content_projecttree) == 0:
            self.updatedisplay("没有获取到任何项目信息！".decode('gbk'))
            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            diag_error_no_project = wx.MessageDialog(None, "没有获取到任何项目信息！".decode('gbk'), '错误'.decode('gbk'),
                                                     wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            diag_error_no_project.ShowModal()
        else:
            for item_projectname in content_projecttree:
                self.listbox_projectname.Append(item_projectname["text"])
        self.updatedisplay("结束获取项目名称".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        diag_finish_project = wx.MessageDialog(None, "获取项目信息完成！".decode('gbk'), '提示'.decode('gbk'),
                                               wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        diag_finish_project.ShowModal()

    def get_phase(self, event):
        self.updatedisplay("开始获取项目阶段".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))

        # 获取图形界面上输入的信息
        username = self.input_username.GetValue()
        password = self.input_password.GetValue()
        project_selected = self.listbox_projectname.GetStringSelection()

        # 登录
        url_login = "http://{}/iauto_acp/login".format(address_web)
        headers_login = {
            'Accept': '*/*',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Content-Length': '32',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Host': '{}'.format(address_web),
            'Origin': 'http://{}'.format(address_web),
            'Referer': 'http://{}/iauto_acp/login.html'.format(address_web),
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
        }
        payload_login = "username={}&password={}".format(username, password)
        get_data = requests.session()
        get_data.post(url_login, data=payload_login, headers=headers_login)

        # 获取所有项目名称
        url_projecttree = "http://{}/iauto_acp/projectAndRound.do/loadProjectTree".format(address_web)
        payload_projecttree = "id=0"
        headers_projecttree = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Content-Length': '4',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Host': '{}'.format(address_web),
            'Origin': 'http://{}'.format(address_web),
            'Referer': 'http://{}/iauto_acp/itmsTestCase.do/testAdmin.view'.format(address_web),
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
        }

        content_projecttree_temp = get_data.post(url_projecttree, data=payload_projecttree,
                                                 headers=headers_projecttree).text
        content_projecttree = json.loads(content_projecttree_temp)
        # 找到选中的项目的ID
        for item_projectname in content_projecttree:
            if item_projectname["text"] == project_selected:
                id_project_selected = item_projectname["id"]
        # 获取所有阶段
        payload_project_selected = "id={}".format(id_project_selected)
        content_project_phase_temp = get_data.post(url_projecttree, data=payload_project_selected,
                                                   headers=headers_projecttree).text
        content_project_phase = json.loads(content_project_phase_temp)
        # show phase
        if len(content_project_phase) == 0:
            self.updatedisplay("没有获取到任何阶段信息！".decode('gbk'))
            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            diag_error_no_project = wx.MessageDialog(None, "没有获取到任何阶段信息！".decode('gbk'), '错误'.decode('gbk'),
                                                     wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            diag_error_no_project.ShowModal()
        else:
            for item_phase in content_project_phase:
                if item_phase["text"] != "项目用例库".decode('gbk'):
                    self.listbox_phase.Append(item_phase["text"])
        self.updatedisplay("结束获取项目阶段！".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        diag_finish_project = wx.MessageDialog(None, "获取项目阶段！完成！".decode('gbk'), '提示'.decode('gbk'),
                                               wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        diag_finish_project.ShowModal()

    def run(self):
        self.updatedisplay("开始获取项目名称".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))

        # 获取图形界面上输入的信息
        username = self.input_username.GetValue()
        password = self.input_password.GetValue()
        project_selected = self.listbox_projectname.GetStringSelection()
        phase_selected = self.listbox_phase.GetStringSelection()
        flag_status_list = ["NP", "FAIL", "BLOCK"]

        # 登录
        url_login = "http://{}/iauto_acp/login".format(address_web)
        headers_login = {
            'Accept': '*/*',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Content-Length': '32',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Host': '{}'.format(address_web),
            'Origin': 'http://{}'.format(address_web),
            'Referer': 'http://{}/iauto_acp/login.html'.format(address_web),
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
        }
        payload_login = "username={}&password={}".format(username, password)
        get_data = requests.session()
        get_data.post(url_login, data=payload_login, headers=headers_login)

        # 获取所有项目名称
        url_projecttree = "http://{}/iauto_acp/projectAndRound.do/loadProjectTree".format(address_web)
        payload_projecttree = "id=0"
        headers_projecttree = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Content-Length': '4',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Host': '{}'.format(address_web),
            'Origin': 'http://{}'.format(address_web),
            'Referer': 'http://{}/iauto_acp/itmsTestCase.do/testAdmin.view'.format(address_web),
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
        }

        content_projecttree_temp = get_data.post(url_projecttree, data=payload_projecttree,
                                                 headers=headers_projecttree).text
        content_projecttree = json.loads(content_projecttree_temp)
        # 找到选中的项目的ID
        for item_projectname in content_projecttree:
            if item_projectname["text"] == project_selected:
                id_project_selected = item_projectname["id"]
        # 找到选中的项目选中的阶段
        # 获取所有阶段
        payload_project_selected = "id={}".format(id_project_selected)
        content_project_phase_temp = get_data.post(url_projecttree, data=payload_project_selected,
                                                   headers=headers_projecttree).text
        content_project_phase = json.loads(content_project_phase_temp)
        for item_phase in content_project_phase:
            if item_phase["text"] == phase_selected:
                id_phase = item_phase["id"]
        # 获取阶段下所有CFG的名称和ID
        payload_config = "id={}".format(id_phase)
        content_config_temp = get_data.post(url_projecttree, data=payload_config,
                                            headers=headers_projecttree).text
        content_config = json.loads(content_config_temp)
        data_dict = {}
        if len(content_config) == 0:
            self.updatedisplay("此阶段下无配置!".decode('gbk'))
            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            diag_error_no_project = wx.MessageDialog(None, "此阶段下无配置！".decode('gbk'), '错误'.decode('gbk'),
                                                     wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            diag_error_no_project.ShowModal()
            self.button_go.Enable()
        else:
            self.updatedisplay("开始获取此阶段配置的用例信息!".decode('gbk'))
            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            for item_config_1 in content_config:
                id_config = item_config_1["id"]
                name_config = item_config_1["text"]
                data_dict["{}".format(id_config)] = {}  # 每个CFG一个dict，后面每个dict写一个sheet页
                data_dict["{}".format(id_config)]["id"] = id_config  # dict中保存CFG的id
                data_dict["{}".format(id_config)]["name"] = name_config  # dict中保存CFG的名称
                data_dict["{}".format(id_config)]["data"] = {}  # dict中data保存往下层级的信息

            for item_config, value_config in data_dict.items():
                data_detail_dict = {}
                parent_id_1 = item_config
                ids_testcase_list = []  # 保存每个配置下最后需要逐个获取详细case信息的列表
                data_return_1 = get_next(get_data, item_config, headers_projecttree, url_projecttree)
                #     return data_return_dict
                self.updatedisplay("开始获取配置下的层级关系-{}!".decode('gbk').format(value_config["name"]))
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                if data_return_1 is None:
                    pass
                else:
                    ids_next_run_1 = []
                    for item_return_1, value_return_1 in data_return_1.items():
                        id_return_1 = value_return_1["id"]
                        testcase_number_1 = value_return_1["casenumber"]
                        if testcase_number_1 != "None":
                            ids_testcase_list.append(id_return_1)
                            data_detail_dict["{}".format(id_return_1)] = value_return_1
                        else:
                            ids_next_run_1.append(id_return_1)
                            data_dict_return_1 = add_item_to_dict(value_return_1, parent_id_1, data_dict)
                            data_dict = data_dict_return_1
                            print(id_return_1)
                    print("1")

                    if len(ids_next_run_1) != 0:
                        ids_next_run_2 = []
                        for item_run_1 in ids_next_run_1:
                            parent_id_2 = item_run_1
                            data_return_2 = get_next(get_data, item_run_1, headers_projecttree, url_projecttree)
                            if data_return_2 is None:
                                pass
                            else:
                                for item_return_2, value_return_2 in data_return_2.items():
                                    id_return_2 = value_return_2["id"]
                                    testcase_number_2 = value_return_2["casenumber"]
                                    if testcase_number_2 != "None":
                                        ids_testcase_list.append(id_return_2)
                                        data_detail_dict["{}".format(id_return_2)] = value_return_2
                                    else:
                                        ids_next_run_2.append(id_return_2)
                                        data_dict_return_2 = add_item_to_dict(value_return_2, parent_id_2, data_dict)
                                        data_dict = data_dict_return_2
                                        print(id_return_2)
                                print("2")

                                if len(ids_next_run_2) != 0:
                                    ids_next_run_3 = []
                                    for item_run_2 in ids_next_run_2:
                                        parent_id_3 = item_run_2
                                        data_return_3 = get_next(get_data, item_run_2, headers_projecttree,
                                                                 url_projecttree)
                                        if data_return_3 is None:
                                            pass
                                        else:
                                            for item_return_3, value_return_3 in data_return_3.items():
                                                id_return_3 = value_return_3["id"]
                                                testcase_number_3 = value_return_3["casenumber"]
                                                if testcase_number_3 != "None":
                                                    ids_testcase_list.append(id_return_3)
                                                    data_detail_dict["{}".format(id_return_3)] = value_return_3
                                                else:
                                                    ids_next_run_3.append(id_return_3)
                                                    data_dict_return_3 = add_item_to_dict(value_return_3, parent_id_3,
                                                                                          data_dict)
                                                    data_dict = data_dict_return_3
                                                    print(id_return_3)
                                            print("3")

                                            if len(ids_next_run_3) != 0:
                                                ids_next_run_4 = []
                                                for item_run_3 in ids_next_run_3:
                                                    parent_id_4 = item_run_3
                                                    data_return_4 = get_next(get_data, item_run_3, headers_projecttree,
                                                                             url_projecttree)
                                                    if data_return_4 is None:
                                                        pass
                                                    else:
                                                        for item_return_4, value_return_4 in data_return_4.items():
                                                            id_return_4 = value_return_4["id"]
                                                            testcase_number_4 = value_return_4["casenumber"]
                                                            if testcase_number_4 != "None":
                                                                ids_testcase_list.append(id_return_4)
                                                                data_detail_dict[
                                                                    "{}".format(id_return_4)] = value_return_4
                                                            else:
                                                                ids_next_run_4.append(id_return_4)
                                                                data_dict_return_4 = add_item_to_dict(value_return_4,
                                                                                                      parent_id_4,
                                                                                                      data_dict)
                                                                data_dict = data_dict_return_4
                                                                print(id_return_4)
                                                        print("3")

                                                        if len(ids_next_run_4) != 0:
                                                            ids_next_run_5 = []
                                                            for item_run_4 in ids_next_run_4:
                                                                parent_id_5 = item_run_4
                                                                data_return_5 = get_next(get_data, item_run_4, headers_projecttree,
                                                                                         url_projecttree)
                                                                if data_return_5 is None:
                                                                    pass
                                                                else:
                                                                    for item_return_5, value_return_5 in data_return_5.items():
                                                                        id_return_5 = value_return_5["id"]
                                                                        testcase_number_5 = value_return_5["casenumber"]
                                                                        if testcase_number_5 != "None":
                                                                            ids_testcase_list.append(id_return_5)
                                                                            data_detail_dict[
                                                                                "{}".format(id_return_5)] = value_return_5
                                                                        else:
                                                                            ids_next_run_5.append(id_return_5)
                                                                            data_dict_return_5 = add_item_to_dict(value_return_5,
                                                                                                                  parent_id_5,
                                                                                                                  data_dict)
                                                                            data_dict = data_dict_return_5
                                                                            print(id_return_5)

                self.updatedisplay("完成获取配置下的层级关系-{}!".decode('gbk').format(value_config["name"]))
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                # 获取每个用例的详细信息
                self.updatedisplay("开始获取每个用例的详细信息-{}!".decode('gbk').format(value_config["name"]))
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                temp_detail = []
                pool_detail = Pool(multiprocessing.cpu_count())
                for id_testcase_all in ids_testcase_list:
                    temp_detail.append(
                        pool_detail.apply_async(get_detail, args=(get_data, id_testcase_all, flag_status_list)))
                pool_detail.close()
                pool_detail.join()
                # return id_testcase_all, bug_id, bug_content, content_bak, status_list, step_list, expect_list, content_list
                for item_return_detail in temp_detail:
                    data_detail_temp = item_return_detail.get()
                    id_testcase = data_detail_temp[0]
                    bug_id = data_detail_temp[1]
                    bug_content = data_detail_temp[2]
                    content_bak = data_detail_temp[3]
                    status_case_list = data_detail_temp[4]
                    step_case_list = data_detail_temp[5]
                    expect_case_list = data_detail_temp[6]
                    content_case_list = data_detail_temp[7]

                    data_detail_dict["{}".format(id_testcase)]["data"]["bug_id"] = bug_id
                    data_detail_dict["{}".format(id_testcase)]["data"]["bug_content"] = bug_content
                    data_detail_dict["{}".format(id_testcase)]["data"]["content_bak"] = content_bak
                    data_detail_dict["{}".format(id_testcase)]["data"]["status_case_list"] = status_case_list
                    data_detail_dict["{}".format(id_testcase)]["data"]["step_case_list"] = step_case_list
                    data_detail_dict["{}".format(id_testcase)]["data"]["expect_case_list"] = expect_case_list
                    data_detail_dict["{}".format(id_testcase)]["data"]["content_case_list"] = content_case_list
                self.updatedisplay("完成获取每个用例的详细信息-{}!".decode('gbk').format(value_config["name"]))
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                print("test")

        self.updatedisplay("完成获取此阶段配置的用例信息!".decode('gbk'))
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        print("END")
        self.button_go.Enable()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = wx.App()
    frame = GetProjectStatus(None)
    frame.Show()
    app.MainLoop()
