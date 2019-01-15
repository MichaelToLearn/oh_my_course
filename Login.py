# coding:utf-8
import json
import pickle

import requests
import os
import re
from io import BytesIO

BASE_DIR = os.path.dirname(__file__)
# LOGIN_URL = 'http://grdms.bit.edu.cn/yjs/login_cas.jsp'
# LOGIN_URL = 'https://login.bit.edu.cn/cas/login?service=https://login.bit.edu.cn/campus-account/shiro-cas'
LOGIN_URL = 'https://login.bit.edu.cn/cas/login?service=http%3A%2F%2Fgrdms.bit.edu.cn%2Fyjs%2Flogin_cas.jsp'
# LOGIN_INDEX_URL = 'https://login.bit.edu.cn/cas/login?service=https://login.bit.edu.cn/campus-account/shiro-cas'
LOGIN_INDEX_URL = LOGIN_URL
# 验证码
CAPTCHA_URL = 'https://login.bit.edu.cn/cas/captcha.html'
NEED_CAPTCHA_URL = 'https://login.bit.edu.cn/cas/needCaptcha.html?username=%s'


class Login(object):

    stuid = ''
    pwd = ''
    cookies = None
    http_session_id = ''

    def __init__(self, stuid, pwd):
        r"""
        初始化
        :param stuid:       学号
        :param pwd:         密码
        """
        self.pwd = pwd
        self.stuid = stuid

    def get_login_param(self):
        r"""
        访问初始登录界面，获取登录参数
        :return:            参数结果，格式如下：
        {
            param: {
                'lt': lt参数,
                'execution': execution参数,
                '_eventId': 事件参数,
                'rmShown': 不知道
            },
            cookies: 请求后所产生的cookie
        }
        """
        r = requests.get(LOGIN_INDEX_URL, verify=False)
        html = r.text
        pattern = 'name="(\S+)"\svalue="(\S+)"'
        search_result = re.findall(pattern, html)
        param = {}
        for item in search_result:
            param[item[0]] = item[1]
        self.cookies = r.cookies
        self.http_session_id = self.get_http_session_id()
        return {
            "param": param,
            "cookies": r.cookies
        }

    def cookies_to_str(self, cookies):
        r"""
        将对象形式的cookie转化成字符串形式
        :param cookies:         对象形式的cookie
        :return:                字符串形式的cookie
        """
        cookie_str = ''
        for index, cookie in enumerate(cookies.items()):
            if index > 0:
                cookie_str = cookie_str + '; '
            cookie_str = cookie_str + "%s=%s" % (cookie[0], cookie[1])
        return cookie_str

    def need_captcha(self):
        r"""
        查询是否需要输入验证码
        :return:            是/否
        """
        r = requests.get(NEED_CAPTCHA_URL % self.stuid, verify=False, headers=self.get_cookie_header())
        return r.json()

    def handle_captcha(self):
        r"""
        处理验证码
        :return:        无
        """
        r = requests.get(CAPTCHA_URL, verify=False, headers=self.get_cookie_header())
        base_dir = os.path.dirname(__file__)
        file_path = os.path.join(base_dir, 'captcha.jpg')
        open(file_path, 'wb').write(r.content)

    def login(self, using_cache=True):
        r"""
        登录
        :return: 登录结果
        {
            "status": 登录是否成功,
            "cookies": 登录成功之后的cookie
        }
        """
        # 判断之前是否保存过登录信息
        if self.logined() and using_cache:
            return {
                'status': True,
                'msg': u'使用之前的缓存登录成功'
            }
        result = self.get_login_param()
        param = result['param']
        param['username'] = self.stuid
        param['password'] = self.pwd
        # 判断是否需要输入验证码
        if self.need_captcha():
            # 需要输入验证码
            self.handle_captcha()
            print(u'请输入验证码：')
            param['captchaResponse'] = input()
        r = requests.post(LOGIN_URL, data=param, verify=False, headers=self.get_cookie_header())
        html = r.text
        msgs = re.findall('id="msg"\sstyle="font-size:20px;color:red;">(.+)</div>', html)
        url_changed = r.url.find('grdms.bit.edu.cn/yjs/login_cas.jsp') != -1
        key_word_contained = html.find(u"top.location = '/yjs/application/main.jsp'") != -1
        if url_changed or key_word_contained:
            print(r.cookies.items())
            # 登录成功
            self.cookies = r.cookies
            # 保存登录信息
            self.save_cookies()
            return {
                'status': True,
                'msg': u'登录成功!'
            }
        else:
            # 登录失败
            if msgs:
                return {
                    'status': False,
                    'msg': msgs[0]
                }
            else:
                return {
                    'status': False,
                    'msg': u'未知错误信息'
                }

    def get_http_session_id(self, prefix=False):
        r"""
        获取档期那的http_session_id
        :return:            http_session_id
        """
        http_session_id = ''
        for item in self.cookies.items():
            if item[0] == 'JSESSIONID':
                http_session_id = item[1]
        if prefix:
            http_session_id = 'ysj1app1~' + http_session_id
        return http_session_id

    def logined(self):
        r"""
        判断是否已经登录过
        :return:    是否
        """
        path = os.path.join(BASE_DIR, 'cookie.txt')
        if os.path.exists(path):
            self.cookies = self.read_cookie()
            return True
        else:
            return False

    def read_cookie(self):
        r"""
        读取cookie
        :return:    无
        """
        path = os.path.join(BASE_DIR, 'cookie.txt')
        cookies = pickle.loads(open(path, 'rb').read())
        return cookies

    def save_cookies(self):
        r"""
        保存cookie
        :return:    无
        """
        path = os.path.join(BASE_DIR, 'cookie.txt')
        open(path, 'wb').write(
            pickle.dumps(self.cookies)
        )

    def get_cookie_header(self):
        r"""
        获取带有cookie的header
        :return:        {
                            "Cookie": ...
                        }
        """
        return {
            "Cookie": self.cookies_to_str(self.cookies)
        }

    def get_header_whole(self):
        r"""
        获取完整的header
        :return:            完整的header
        """
        headers = self.get_cookie_header()
        headers_other = {
            'Accept': 'image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*',
            'Referer': 'http://grdms.bit.edu.cn/yjs/yanyuan/py/pyjxjh.do?method=queryListing',
            'Accept-Language': 'zh-CN',
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; doyo 2.6.1)',
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept-Encoding': 'gzip, deflate',
            'Content-Length': '668',
            'Host': 'grdms.bit.edu.cn',
            'Connection': 'Keep-Alive',
            'Pragma': 'no-cache'
        }
        headers = dict(headers, **headers_other)
        return headers