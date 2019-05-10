#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import requests
import re
from bs4 import BeautifulSoup
import xlsxwriter
import os
import copy
import time
from threading import Thread
import wx
from multiprocessing import Pool
import multiprocessing
import json

'''20190422:Version-1.0:Initial Version'''

num_year = "2019"
ver = "1.0"
address_web_dict = {u"内网": "100.2.39.222", u"外网": "172.31.2.125"}


def get_next(get_data, id_sub, headers_sub, url_sub):
    payload_next_sub = "id={}".format(id_sub)
    get_page = get_data.post(url_sub, headers=headers_sub, data=payload_next_sub)
    data_page = json.loads(get_page.text)
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


def get_next_detail(get_data, id_sub, headers_sub, url_sub):
    payload_next_sub = "id={}".format(id_sub)
    get_page = get_data.post(url_sub, headers=headers_sub, data=payload_next_sub)
    data_page = json.loads(get_page.text)
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
            if len(text_data.split(";")) == 1:
                data_return_dict["{}".format(id_data)]["name"] = text_data
            else:
                data_return_dict["{}".format(id_data)]["name"] = text_data.split(";")[1]
            data_return_dict["{}".format(id_data)]["casenumber"] = casenumber_data
            data_return_dict["{}".format(id_data)]["data"] = {}
            data_return_dict["{}".format(id_data)]["parentid"] = id_sub

    return data_return_dict


def add_item_to_dict(value_return, parent_id, dict_sub):
    dict_value_sub = dict_sub
    for key_dict, value_dict in dict_value_sub.items():
        if isinstance(value_dict, dict) and "id" in value_dict:
            if value_dict["id"] == parent_id:
                value_dict["data"]["{}".format(value_return["id"])] = value_return
                return dict_value_sub
            else:
                if isinstance(value_dict["data"], dict):
                    add_item_to_dict(value_return, parent_id, value_dict["data"])
    return dict_value_sub


def add_level(get_data, headers_projecttree, url_projecttree, data_dict, data_detail_dict, ids_next_run):
    ids_next_run_list_temp = []
    ids_testcase_list_temp = []
    for item_run in ids_next_run:
        parent_id = item_run
        data_return = get_next_detail(get_data, item_run, headers_projecttree, url_projecttree)
        if data_return is None:
            pass
        else:
            for item_return, value_return in data_return.items():
                id_return = value_return["id"]
                testcase_number = value_return["casenumber"]
                if testcase_number != "None":
                    ids_testcase_list_temp.append(id_return)
                    data_detail_dict["{}".format(id_return)] = value_return
                else:
                    ids_next_run_list_temp.append(id_return)
                    data_dict_return = add_item_to_dict(value_return, parent_id, data_dict)
                    data_dict = data_dict_return
                    # print(id_return)
    return ids_next_run_list_temp, ids_testcase_list_temp, data_dict, data_detail_dict


def get_detail(get_data, id_testcase_all, flag_status_list, address_web):
    get_data_sub = get_data
    url_testcase = "http://{}/iauto_acp/itmsTestCaseN.do/projectConfigTestCaseInfo.view".format(address_web)
    headers_detail = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
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
    print(id_testcase_all)
    print(page_testcase_temp.status_code)
    page_testcase = BeautifulSoup(page_testcase_temp.text, "html.parser")
    # bug ID
    bug_id = page_testcase.find("td", text="Bug Id:").parent.find("a").get_text()
    # BUG内容
    bug_content = page_testcase.find("td", text="Bug Content:").parent.find("a").get_text()
    # 备注
    content_bak = page_testcase.find("td", text="备注(content):").parent.find("a").get_text()
    produce_temp = re.search(r'var procedureList = (\[\{.*?\}\]);', page_testcase_temp.text).groups()[0]
    produce = json.loads(produce_temp)
    step_list = []
    expect_list = []
    status_list = []
    content_list = []
    remark_step_list = []
    if len(produce) != 0:
        for item_step in produce:
            status_step = item_step["result"]
            if status_step in flag_status_list:
                status_list.append(status_step)
                step_list.append(item_step["testProcedure"])
                expect_list.append(item_step["testExpect"])
                content_list.append(item_step["remark"])

    get_data_sub.close()
    return id_testcase_all, bug_id, bug_content, content_bak, status_list, step_list, expect_list, content_list


# data_to_write(value_case_write, parent_id_write, data_dict[item_config]["data"])
def data_to_write(parent_id_write, dict_sub, list_return):
    dict_value_sub = dict_sub
    # value_case_write = value_case_write_sub
    list_to_return = copy.deepcopy(list_return)
    for key, val in dict_value_sub.items():
        if isinstance(val, dict) and "id" in val and "name" in val:
            if parent_id_write == key:
                list_to_return.append(val["name"])
                return val, list_to_return
            elif isinstance(val["data"], dict):
                list_to_return.append(val["name"])
                result_tuple = data_to_write(parent_id_write, val["data"], list_to_return)
                if result_tuple:
                    return result_tuple[0], result_tuple[1]
                else:
                    if list_to_return:
                        list_to_return = list_to_return[0:len(list_to_return) - 1]


class GetProjectStatus(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"获取ITMS系统上某个项目某个阶段的所有配置的NP或BLOCK或的项用例项目-{}".format(ver),
                          pos=wx.DefaultPosition, size=wx.Size(504, 821),
                          style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_APPWORKSPACE))

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel1 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        self.m_panel1.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))

        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        bSizer31 = wx.BoxSizer(wx.VERTICAL)

        self.text_title14 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 1.请选择是内网还是外网！", wx.DefaultPosition,
                                          wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title14.Wrap(-1)

        self.text_title14.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title14.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title14.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer31.Add(self.text_title14, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer31, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer191 = wx.BoxSizer(wx.VERTICAL)

        listbox_placesChoices = [u"内网", u"外网"]
        self.listbox_places = wx.ListBox(self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                         listbox_placesChoices, wx.LB_ALWAYS_SB | wx.LB_SINGLE)
        bSizer191.Add(self.listbox_places, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

        bSizer10.Add(bSizer191, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_title1 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 2.请输入ITMS系统的用户名和密码！\n然后点击下方按钮！",
                                         wx.DefaultPosition, wx.DefaultSize,
                                         wx.ALIGN_CENTER_HORIZONTAL | wx.ST_NO_AUTORESIZE)
        self.text_title1.Wrap(-1)

        self.text_title1.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title1.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title1.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer3.Add(self.text_title1, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

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

        self.text_title11 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 3.请选择要分析的项目！然后点击下方按钮！", wx.DefaultPosition,
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

        self.text_title13 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 4：请选择阶段！", wx.DefaultPosition,
                                          wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title13.Wrap(-1)

        self.text_title13.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title13.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title13.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer13.Add(self.text_title13, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer12.Add(bSizer13, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer15 = wx.BoxSizer(wx.VERTICAL)

        listbox_phaseChoices = []
        self.listbox_phase = wx.ListBox(self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                        listbox_phaseChoices, 0)
        bSizer15.Add(self.listbox_phase, 1, wx.ALL | wx.EXPAND, 5)

        bSizer12.Add(bSizer15, 1, wx.EXPAND, 5)

        bSizer10.Add(bSizer12, 0, wx.EXPAND, 5)

        bSizer121 = wx.BoxSizer(wx.VERTICAL)

        bSizer131 = wx.BoxSizer(wx.VERTICAL)

        self.text_title131 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 5:请选择要获取的异常用例的状态！可以多选！", wx.DefaultPosition,
                                           wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title131.Wrap(-1)

        self.text_title131.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title131.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title131.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer131.Add(self.text_title131, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer121.Add(bSizer131, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer151 = wx.BoxSizer(wx.VERTICAL)

        listbox_statusChoices = [u"NP", u"BLOCK", u"FAIL"]
        self.listbox_status = wx.ListBox(self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                         listbox_statusChoices, wx.LB_ALWAYS_SB | wx.LB_MULTIPLE)
        bSizer151.Add(self.listbox_status, 1, wx.EXPAND | wx.ALL, 5)

        bSizer121.Add(bSizer151, 1, wx.EXPAND, 5)

        bSizer10.Add(bSizer121, 1, wx.EXPAND, 5)

        bSizer21 = wx.BoxSizer(wx.VERTICAL)

        bSizer211 = wx.BoxSizer(wx.VERTICAL)

        self.text_title12 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 6.请点击GO开始导出！或者点击EXIT退出程序！",
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

        # self._thread = Thread(target=self.run, args=())
        # self._thread.daemon = True

    def close(self, event):
        self.Close()

    def newthread(self):
        Thread(target=self.run_all).start()

    def onbutton(self, event):
        # self._thread.start()
        # self.started = True
        # self.button_go = event.GetEventObject()
        self.newthread()
        self.button_go.Disable()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.textctrl_display.AppendText("完成{}%".format(t))
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.textctrl_display.AppendText("%s" % t)
        self.textctrl_display.AppendText(os.linesep)

    def get_projectname(self, event):
        self.updatedisplay("开始获取项目名称")
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))

        # 获取图形界面上输入的信息
        # 内网还是外网
        places = self.listbox_places.GetStringSelection()
        # 用户名
        username = self.input_username.GetValue()
        # 密码
        password = self.input_password.GetValue()

        # 登录
        if len(places) == 0:
            self.updatedisplay("没有选择内网还是外网！")
            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            diag_error = wx.MessageDialog(None, "没有选择内网还是外网！", '错误',
                                          wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            diag_error.ShowModal()
        else:
            if len(username) == 0 or len(password) == 0:
                self.updatedisplay("没有输入用户名或者密码！")
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                diag_error = wx.MessageDialog(None, "没有输入用户名或者密码！", '错误',
                                              wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                diag_error.ShowModal()
            else:
                address_web = address_web_dict[places.encode('unicode_escape').decode('unicode_escape')]
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
                    self.updatedisplay("没有获取到任何项目信息！")
                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                    diag_error_no_project = wx.MessageDialog(None, "没有获取到任何项目信息！", '错误',
                                                             wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                    diag_error_no_project.ShowModal()
                else:
                    for item_projectname in content_projecttree:
                        self.listbox_projectname.Append(item_projectname["text"])
                self.updatedisplay("结束获取项目名称")
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                diag_finish_project = wx.MessageDialog(None, "获取项目信息完成！", '提示',
                                                       wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
                diag_finish_project.ShowModal()

    def get_phase(self, event):
        self.updatedisplay("开始获取项目阶段")
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        self.listbox_phase.Clear()
        # 获取图形界面上输入的信息
        # 内网还是外网
        places = self.listbox_places.GetStringSelection()
        # 用户名
        username = self.input_username.GetValue()
        # 密码
        password = self.input_password.GetValue()
        # 选择的项目
        project_selected = self.listbox_projectname.GetStringSelection()

        # 登录
        if len(places) == 0:
            self.updatedisplay("没有选择内网还是外网！")
            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            diag_error = wx.MessageDialog(None, "没有选择内网还是外网！", '错误',
                                          wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            diag_error.ShowModal()
        else:
            if len(username) == 0 or len(password) == 0:
                self.updatedisplay("没有输入用户名或者密码！")
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                diag_error = wx.MessageDialog(None, "没有输入用户名或者密码！", '错误',
                                              wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                diag_error.ShowModal()
            else:
                if len(project_selected) == 0:
                    self.updatedisplay("没有选择项目名称！")
                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                    diag_error = wx.MessageDialog(None, "没有选择项目名称！", '错误',
                                                  wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                    diag_error.ShowModal()
                else:
                    address_web = address_web_dict[places.encode('unicode_escape').decode('unicode_escape')]
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
                        self.updatedisplay("没有获取到任何阶段信息！")
                        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                        diag_error_no_project = wx.MessageDialog(None, "没有获取到任何阶段信息！", '错误',
                                                                 wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                        diag_error_no_project.ShowModal()
                    else:
                        for item_phase in content_project_phase:
                            if item_phase["text"] != "项目用例库":
                                self.listbox_phase.Append(item_phase["text"])
                    self.updatedisplay("结束获取项目阶段！")
                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                    diag_finish_project = wx.MessageDialog(None, "获取项目阶段！完成！", '提示',
                                                           wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
                    diag_finish_project.ShowModal()

    def run_all(self):
        self.updatedisplay("开始获取项目名称")
        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))

        # 获取图形界面上输入的信息
        # 内网还是外网
        places = self.listbox_places.GetStringSelection()
        # 用户名
        username = self.input_username.GetValue()
        # 密码
        password = self.input_password.GetValue()
        # 选择的项目
        project_selected = self.listbox_projectname.GetStringSelection()
        # 选择的阶段
        phase_selected = self.listbox_phase.GetStringSelection()
        # 选择的异常状态
        flag_status_list_temp = self.listbox_status.GetSelections()
        flag_status_list = []
        for item in flag_status_list_temp:
            flag_status_list.append(self.listbox_status.GetString(item))
        # 登录
        if len(places) == 0:
            self.updatedisplay("没有选择内网还是外网！请退出程序，然后重新打开！")
            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            diag_error = wx.MessageDialog(None, "没有选择内网还是外网！请退出程序，然后重新打开！", '错误',
                                          wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            diag_error.ShowModal()
            self.button_go.Enable()
        else:
            if len(username) == 0 or len(password) == 0:
                self.updatedisplay("没有输入用户名或者密码！请退出程序，然后重新打开！")
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                diag_error = wx.MessageDialog(None, "没有输入用户名或者密码！请退出程序，然后重新打开！", '错误',
                                              wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                diag_error.ShowModal()
                self.button_go.Enable()
            else:
                if len(project_selected) == 0:
                    self.updatedisplay("没有选择项目名称！请退出程序，然后重新打开！")
                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                    diag_error = wx.MessageDialog(None, "没有选择项目名称！请退出程序，然后重新打开！", '错误',
                                                  wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                    diag_error.ShowModal()
                    self.button_go.Enable()
                else:
                    if len(phase_selected) == 0:
                        self.updatedisplay("没有选择项目阶段！请退出程序，然后重新打开！")
                        self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                        diag_error = wx.MessageDialog(None, "没有选择项目阶段！请退出程序，然后重新打开！", '错误',
                                                      wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                        diag_error.ShowModal()
                        self.button_go.Enable()
                    else:
                        if len(flag_status_list) == 0:
                            self.updatedisplay("没有选择要获取的异常状态！请退出程序，然后重新打开！")
                            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                            diag_error = wx.MessageDialog(None, "没有选择要获取的异常状态！请退出程序，然后重新打开！",
                                                          '错误',
                                                          wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                            diag_error.ShowModal()
                            self.button_go.Enable()
                        else:
                            address_web = address_web_dict[places.encode('unicode_escape').decode('unicode_escape')]
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
                            url_projecttree = "http://{}/iauto_acp/projectAndRound.do/loadProjectTree".format(
                                address_web)
                            payload_projecttree = "id=0"
                            headers_projecttree = {
                                'Accept': 'application/json, text/javascript, */*; q=0.01',
                                'Accept-Encoding': 'gzip, deflate',
                                'Accept-Language': 'zh-CN,zh;q=0.9',
                                'Connection': 'keep-alive',
                                'Content-Length': '37',
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
                                    project_name = item_projectname["text"]
                            # 找到选中的项目选中的阶段
                            # 获取所有阶段
                            payload_project_selected = "id={}".format(id_project_selected)
                            content_project_phase_temp = get_data.post(url_projecttree, data=payload_project_selected,
                                                                       headers=headers_projecttree).text
                            content_project_phase = json.loads(content_project_phase_temp)
                            for item_phase in content_project_phase:
                                if item_phase["text"] == phase_selected:
                                    id_phase = item_phase["id"]
                                    phase_name = item_phase["text"]
                            # 获取阶段下所有CFG的名称和ID
                            payload_config = "id={}".format(id_phase)
                            content_config_temp = get_data.post(url_projecttree, data=payload_config,
                                                                headers=headers_projecttree).text
                            content_config = json.loads(content_config_temp)
                            data_dict = {}

                            if len(content_config) == 0:
                                self.updatedisplay("此阶段下无配置!")
                                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                diag_error_no_project = wx.MessageDialog(None, "此阶段下无配置！",
                                                                         '错误',
                                                                         wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                                diag_error_no_project.ShowModal()
                                self.button_go.Enable()
                            else:
                                self.updatedisplay("开始获取此阶段配置的用例信息!")
                                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                # 创建保存输出结果的excel表格
                                TitleItem = ['层级1', '层级2', '层级3',
                                             '层级4',
                                             '用例名称', '用例编号', '用例BUG ID',
                                             '用例BUG描述',
                                             '用例备注', '测试步骤', '预期结果',
                                             '测试结果状态','测试步骤备注']
                                timestamp = time.strftime('%Y%m%d', time.localtime())
                                WorkBook = xlsxwriter.Workbook("{}项目{}阶段配置下异常用例信息获取结果-{}.xlsx".format(project_name, phase_name, timestamp))
                                formatOne = WorkBook.add_format()
                                formatOne.set_border(1)
                                for item_config_1 in content_config:
                                    id_config = item_config_1["id"]
                                    name_config = item_config_1["text"]
                                    data_dict["{}".format(id_config)] = {}  # 每个CFG一个dict，后面每个dict写一个sheet页
                                    data_dict["{}".format(id_config)]["id"] = id_config  # dict中保存CFG的id
                                    data_dict["{}".format(id_config)]["data"] = {}  # dict中data保存往下层级的信息
                                    data_dict["{}".format(id_config)]["name"] = name_config  # dict中保存CFG的名称

                                for item_config, value_config in data_dict.items():
                                    # 创建以CFG名称命名的sheet页面
                                    Sheet = WorkBook.add_worksheet('{}'.format(value_config["name"]))
                                    # 在sheet中写入标题
                                    for i in range(0, len(TitleItem)):
                                        Sheet.write(0, i, TitleItem[i], formatOne)
                                    # 设置列宽
                                    Sheet.set_column('A:D', 14)
                                    Sheet.set_column('E:E', 35)
                                    Sheet.set_column('F:F', 25)
                                    Sheet.set_column('G:I', 20)
                                    Sheet.set_column('J:K', 30)
                                    Sheet.set_column('L:M', 13)

                                    data_detail_dict = {}
                                    parent_id_1 = item_config
                                    ids_testcase_list = []  # 保存每个配置下最后需要逐个获取详细case信息的列表
                                    data_return_1 = get_next(get_data, item_config, headers_projecttree,
                                                             url_projecttree)
                                    #     return data_return_dict
                                    self.updatedisplay("开始获取配置下的层级关系-{}!".format(value_config["name"]))
                                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                    if data_return_1 is None:
                                        pass
                                    else:
                                        ids_next_run = []
                                        for item_return_1, value_return_1 in data_return_1.items():
                                            id_return_1 = value_return_1["id"]
                                            testcase_number_1 = value_return_1["casenumber"]
                                            if testcase_number_1 != "None":
                                                ids_testcase_list.append(id_return_1)
                                                data_detail_dict["{}".format(id_return_1)] = value_return_1
                                            else:
                                                ids_next_run.append(id_return_1)
                                                data_dict_return_1 = add_item_to_dict(value_return_1, parent_id_1,
                                                                                      data_dict)
                                                data_dict = data_dict_return_1
                                                print(id_return_1)
                                        while len(ids_next_run) != 0:
                                            print(ids_next_run)
                                            return_add_level = add_level(get_data, headers_projecttree, url_projecttree,
                                                                         data_dict, data_detail_dict, ids_next_run)
                                            # return ids_next_run_list_temp, ids_testcase_list_temp, data_dict, data_detail_dict
                                            ids_next_run = return_add_level[0]
                                            ids_testcase_list.extend(return_add_level[1])
                                            data_dict = return_add_level[2]
                                            data_detail_dict = return_add_level[3]

                                    self.updatedisplay("完成获取配置下的层级关系-{}!".format(value_config["name"]))
                                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                    # 获取每个用例的详细信息
                                    self.updatedisplay("开始获取每个用例的详细信息-{}!".format(value_config["name"]))
                                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                    temp_detail = []
                                    pool_detail = Pool(multiprocessing.cpu_count())
                                    for id_testcase_all in ids_testcase_list:
                                        temp_detail.append(
                                            pool_detail.apply_async(get_detail, args=(
                                                get_data, id_testcase_all, flag_status_list, address_web)))
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

                                        # data_detail_dict["{}".format(id_testcase)]["data"]["id"] = id_testcase
                                        data_detail_dict["{}".format(id_testcase)]["data"]["bug_id"] = bug_id
                                        data_detail_dict["{}".format(id_testcase)]["data"]["bug_content"] = bug_content
                                        data_detail_dict["{}".format(id_testcase)]["data"]["content_bak"] = content_bak
                                        data_detail_dict["{}".format(id_testcase)]["data"][
                                            "status_case_list"] = status_case_list
                                        data_detail_dict["{}".format(id_testcase)]["data"][
                                            "step_case_list"] = step_case_list
                                        data_detail_dict["{}".format(id_testcase)]["data"][
                                            "expect_case_list"] = expect_case_list
                                        data_detail_dict["{}".format(id_testcase)]["data"][
                                            "content_case_list"] = content_case_list
                                        data_detail_dict["{}".format(id_testcase)]["data"]["data"] = "data"

                                    self.updatedisplay("完成获取每个用例的详细信息-{}!".format(value_config["name"]))
                                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                    # 去除获取到的用例中，没有出现异常的
                                    data_detail_dict_write = {}
                                    for item_case, value_case in data_detail_dict.items():
                                        status_case = value_case["data"]["status_case_list"]
                                        if len(status_case) != 0:
                                            data_detail_dict_write["{}".format(item_case)] = value_case

                                    self.updatedisplay("开始将{}信息输出到excel中!".format(value_config["name"]))
                                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                    list_to_write = []
                                    for item_case_write, value_case_write in data_detail_dict_write.items():
                                        parent_id_write = value_case_write["parentid"]
                                        list_return = []
                                        # list_return_write = data_to_write(value_case_write, parent_id_write, data_dict[item_config]["data"], list_return)
                                        list_return_write = data_to_write(parent_id_write,
                                                                          data_dict[item_config]["data"], list_return)
                                        list_return_write_temp = list_return_write[1]
                                        for count in range(0, 4 - len(list_return_write[1])):
                                            list_return_write_temp.append("")
                                        list_return_write_temp.append(value_case_write)
                                        list_to_write.append(list_return_write_temp)
                                    # 写到sheet中
                                    baseline_to_write = 1
                                    for index_write, item_write in enumerate(list_to_write):
                                        # 将某个用例中共用的几列合并单元格,然后写入信息
                                        len_to_merge = len(item_write[4]["data"]["status_case_list"]) - 1
                                        # 如果只有一行，则不需要合并单元格
                                        if len_to_merge == 0:
                                            # 写入层级1
                                            Sheet.write(baseline_to_write + index_write, 0,
                                                        item_write[0], formatOne)
                                            # 写入层级2
                                            Sheet.write(baseline_to_write + index_write, 1,
                                                        item_write[1], formatOne)
                                            # 写入层级3
                                            Sheet.write(baseline_to_write + index_write, 2,
                                                        item_write[2], formatOne)
                                            # 写入层级4
                                            Sheet.write(baseline_to_write + index_write, 3,
                                                        item_write[3], formatOne)
                                            # 写入用例名称
                                            Sheet.write(baseline_to_write + index_write, 4,
                                                        item_write[4]["name"], formatOne)
                                            # 写入用例编号
                                            Sheet.write(baseline_to_write + index_write, 5,
                                                        item_write[4]["casenumber"], formatOne)
                                            # 写入BUG ID
                                            Sheet.write(baseline_to_write + index_write, 6,
                                                        item_write[4]["data"]["bug_id"], formatOne)
                                            # 写入BUG 描述
                                            Sheet.write(baseline_to_write + index_write, 7,
                                                        item_write[4]["data"]["bug_content"], formatOne)
                                            # 写入用例备注
                                            Sheet.write(baseline_to_write + index_write, 8,
                                                        item_write[4]["data"]["content_bak"], formatOne)
                                        else:
                                            # 写入层级1
                                            Sheet.merge_range(baseline_to_write + index_write, 0,
                                                              baseline_to_write + index_write + len_to_merge, 0,
                                                              item_write[0], formatOne)
                                            # 写入层级2
                                            Sheet.merge_range(baseline_to_write + index_write, 1,
                                                              baseline_to_write + index_write + len_to_merge, 1,
                                                              item_write[1], formatOne)
                                            # 写入层级3
                                            Sheet.merge_range(baseline_to_write + index_write, 2,
                                                              baseline_to_write + index_write + len_to_merge, 2,
                                                              item_write[2], formatOne)
                                            # 写入层级4
                                            Sheet.merge_range(baseline_to_write + index_write, 3,
                                                              baseline_to_write + index_write + len_to_merge, 3,
                                                              item_write[3], formatOne)
                                            # 写入用例名称
                                            Sheet.merge_range(baseline_to_write + index_write, 4,
                                                              baseline_to_write + index_write + len_to_merge, 4,
                                                              item_write[4]["name"], formatOne)
                                            # 写入用例编号
                                            Sheet.merge_range(baseline_to_write + index_write, 5,
                                                              baseline_to_write + index_write + len_to_merge, 5,
                                                              item_write[4]["casenumber"], formatOne)
                                            # 写入BUG ID
                                            Sheet.merge_range(baseline_to_write + index_write, 6,
                                                              baseline_to_write + index_write + len_to_merge, 6,
                                                              item_write[4]["data"]["bug_id"], formatOne)
                                            # 写入BUG 描述
                                            Sheet.merge_range(baseline_to_write + index_write, 7,
                                                              baseline_to_write + index_write + len_to_merge, 7,
                                                              item_write[4]["data"]["bug_content"], formatOne)
                                            # 写入用例备注
                                            Sheet.merge_range(baseline_to_write + index_write, 8,
                                                              baseline_to_write + index_write + len_to_merge, 8,
                                                              item_write[4]["data"]["content_bak"], formatOne)
                                        # 写入测试步骤、预期结果、测试结果状态、测试步骤备注
                                        for index_status, item_status in enumerate(
                                                item_write[4]["data"]["status_case_list"]):
                                            # baseline_to_write = baseline_to_write + index_status
                                            Sheet.write(baseline_to_write + index_write + index_status, 9,
                                                        item_write[4]["data"]["step_case_list"][index_status],
                                                        formatOne)
                                            Sheet.write(baseline_to_write + index_write + index_status, 10,
                                                        item_write[4]["data"]["expect_case_list"][index_status],
                                                        formatOne)
                                            Sheet.write(baseline_to_write + index_write + index_status, 11,
                                                        item_write[4]["data"]["status_case_list"][index_status],
                                                        formatOne)
                                            Sheet.write(baseline_to_write + index_write + index_status, 12,
                                                        item_write[4]["data"]["content_case_list"][index_status],
                                                        formatOne)
                                        baseline_to_write = baseline_to_write + len_to_merge
                                    self.updatedisplay("结束将{}的信息输出到excel中!".format(value_config["name"]))
                                    self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                self.updatedisplay("完成获取此阶段配置的用例信息,请点击EXIT退出程序!")
                                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                                WorkBook.close()
                            print("END ALL!")
                            self.button_go.Enable()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = wx.App()
    frame = GetProjectStatus(None)
    frame.Show()
    app.MainLoop()
