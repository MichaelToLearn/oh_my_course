# coding:utf-8
import _thread
import json
import os
import pickle
import re
import threading
import time
from urllib.parse import quote

import requests
from Login import Login
import xlrd
import sys

class Course(Login):

    # 基本信息
    basic_info = {}
    # 课程数据库
    database = {}

    def test(self):
        header = {
            "Cookie": self.cookies_to_str(self.cookies)
        }
        r = requests.get('http://grdms.bit.edu.cn/yjs/application/page_main.jsp?subsys=', cookies=self.cookies, verify=False, headers=header)

    @staticmethod
    def read_account():
        r"""
        读取账号密码
        :return:    {
                        'username': 用户名,
                        'password': 密码
                    }
        """
        return Course.read_courses('', account_only=True)

    def read_courses(self, account_only=False):
        r"""
        读取本地的课程excel表格
        :return:    课程数据，格式如下：
        [
            {
                "class_name": 教学班名称,
                "code": 课程代码,
                "name": 课程中文名称
            },
            ...
        ]
        """
        base_dir = os.path.dirname(__file__)
        file_path = os.path.join(base_dir, u'抢课模板.xlsx')
        if not os.path.exists(file_path):
            print(u'当前目录找不到“抢课模板.xlsx”！')
        else:
            workbook = xlrd.open_workbook(file_path)
            sheet = workbook.sheet_by_index(0)
            rows = sheet.nrows
            data = []
            # 是否是读取账户
            if account_only:
                username = sheet.cell(1, 0).value
                password = sheet.cell(1, 1).value
                currentYear = sheet.cell(1, 2).value
                currentTerm = sheet.cell(1, 3).value
                if type(username) == float:
                    username = str(int(username))
                if type(password) == float:
                    password = str(int(password))
                return {
                    'username': username,
                    'password': password,
                    'currentYear': currentYear,
                    'currentTerm': currentTerm
                }
            for row in range(3, sheet.nrows):
                # 每一行
                row_data = sheet.row_values(row)
                course_class_name = row_data[0]
                code = row_data[1]
                if type(course_class_name) is float:
                    # 说明纯数字，但是excel又把它识别成浮点数了
                    # 此时应该先转换成整数，然后再转换成字符串
                    course_class_name = str(int(course_class_name))
                if type(code) is float:
                    # 说明纯数字，但是excel又把它识别成浮点数了
                    # 此时应该先转换成整数，然后再转换成字符串
                    code = str(int(code))
                data.append({
                    "class_name": course_class_name,
                    "code": code,
                    "name": row_data[2]
                })
            return data

    def get_basic_info(self):
        r"""
        获取基本信息
        :return:        基本信息，格式如下：
        {
            'http_session_id': 当前session的id,
            'script_session_id': 当前脚本的session id,
            'stuid': 学号,
            'student_id': 学生id（为数据库中的独立id）,
            'student_type': 学生类别（如统招统分研究生）,
            'grade': 年级,
            'education': 学历(硕士、博士等),

            'year': 开课学年,
            'term': 开课学期,
            'lb': 类别,
            'pyfaId': 应该就是本学生的培养方案id
        }
        """
        course_index_url = 'http://grdms.bit.edu.cn/yjs/yanyuan/py/pyjxjh.do?method=stdSelectCourseEntry'
        try:
            r = requests.get(course_index_url, headers=self.get_cookie_header())
            self.cookies = r.cookies
            if r.text.find(u'不在选课时间范围') != -1 or r.text.find(u'暂时无法选课') != -1:
                # 说明还没开始选课
                return {
                    "status": False,
                    'msg': u'尚未开始选课'
                }
            else:
                pattern = """name="(\S+)"\svalue=['"](\S+)['"]"""
                result = re.findall(pattern, r.text)
                if not result:
                    return {
                        'status': False
                    }
                obj = {
                    'year': result[0][1],
                    'term': result[1][1],
                    'student_id': result[2][1],
                    'stuid': result[3][1],
                    'lb': result[4][1],
                    'education': result[5][1],
                    'student_type': result[6][1],
                    'grade': result[7][1],
                    'http_session_id': self.get_http_session_id(prefix=True),
                    'script_session_id': self.get_script_session_id(),
                    'pyfaId': result[10][1],
                    'ldfs': result[8][1]
                }
                self.basic_info = obj
                return {
                    'status': True,
                    'obj': obj
                }
        except:
            time.sleep(0.5)
            return {
                "status": False,
                'msg': u'请求次数过多，歇0.5秒'
            }

    def get_script_session_id(self):
        r"""
        获取当前的脚本id
        :return:            脚本id
        """
        return '3A9C853C62112D4DA60AD8D3126408C0'

    def read_database(self):
        r"""
        读取课程数据库
        :return:
        """
        base_dir = os.path.dirname(__file__)
        path = os.path.join(base_dir, 'course_database.txt')
        if os.path.exists(path):
            # 说明课程数据保存过
            self.database = json.loads(open(path, 'r').read())
        else:
            # 说明课程数据没有保存过，直接赋为空值
            self.database = {}

    def push_database(self, course_code, course_class_name, info):
        r"""
        往数据库中添加信息
        :return:
        """
        self.database['%s#%s' % (course_code, course_class_name)] = info

    def pop_database(self, course_code, course_class_name):
        key_name = "%s#%s" % (course_code, course_class_name)
        if key_name in self.database:
            return self.database[key_name]
        else:
            return None

    def write_database(self):
        r"""
        写数据库到文件
        :return:
        """
        base_dir = os.path.dirname(__file__)
        path = os.path.join(base_dir, 'course_database.txt')
        open(path, 'w+').write(json.dumps(self.database))

    def get_course_info(self, course_code, course_class_name, using_online=False):
        r"""
        获取要选的课程的信息
        每次获取之后，要把数据都存储在本地，下一次可以直接获取了，不需要请求网络了.
        所谓数据库，就是一个字典，字典的key是“代码#班级名称”，根据key查询
        :param course_code:             课程代码
        :param course_class_name:       课程班名称
        :return:                        课程信息，格式如下：
        {
            "id": 课程id,
            "code": 课程代码,
            "name": 课程名称,
            "type": 课程类别（如专业必修课）,
            "property": 课程性质（如专业必修课）,
            "hour": 学时,
            "credit": 学分
        }
        """
        # 查看当前课程是否已经被保存到本地
        local_result = self.pop_database(course_code, course_class_name)
        if local_result is not None:
            return {
                'status': True,
                'obj': local_result
            }

        # 从网络上获取课程信息
        if using_online:
            return self.get_course_info_online(course_code, course_class_name)
        # 搜索课程
        course_class_name_encoded = self.encode_chinese(course_class_name)
        if self.currentTerm == u'第一学期':
            post_data = """criteria=KKXN+like+%27%25""" + self.currentYear + """%25%27+escape+%27%5C%27+and+KKXQ+like+%27%25%B5%DA%D2%BB%D1%A7%C6%DA%25%27+escape+%27%5C%27+and+JXBMC+like+%27%25""" + course_class_name_encoded+ """%25%27+escape+%27%5C%27+and+KCDM+like+%27%25""" + course_code + """%25%27+escape+%27%5C%27&vosql=&parameters=&maxPageItems=20&maxIndexPages=5&newID=&defName=pyjxjhStuQuery&defaultcriteria=KKXN+like+%27%25""" + self.currentYear + """%25%27+escape+%27%5C%27+and+KKXQ+like+%27%25%B5%DA%D2%BB%D1%A7%C6%DA%25%27+escape+%27%5C%27+and+JXBMC+like+%27%25""" + course_class_name_encoded+ """%25%27+escape+%27%5C%27+and+KCDM+like+%27%25""" + course_code + """%25%27+escape+%27%5C%27&orderby2=+kkxn%2Ckkxq%2Ckcdm%2Cjxbmc+asc+&seq2=&hiddenColumns=null&pager.offset=0&maxPageItemsSelect=20&maxPageItems2=20"""
        else:
            post_data = """criteria=KKXN+like+%27%25""" + self.currentYear + """%25%27+escape+%27%5C%27+and+KKXQ+like+%27%25%B5%DA%B6%FE%D1%A7%C6%DA%25%27+escape+%27%5C%27+and+JXBMC+like+%27%25""" + course_class_name_encoded+ """%25%27+escape+%27%5C%27+and+KCDM+like+%27%25""" + course_code + """%25%27+escape+%27%5C%27&vosql=&parameters=&maxPageItems=20&maxIndexPages=5&newID=&defName=pyjxjhStuQuery&defaultcriteria=KKXN+like+%27%25""" + self.currentYear + """%25%27+escape+%27%5C%27+and+KKXQ+like+%27%25%B5%DA%B6%FE%D1%A7%C6%DA%25%27+escape+%27%5C%27+and+JXBMC+like+%27%25""" + course_class_name_encoded+ """%25%27+escape+%27%5C%27+and+KCDM+like+%27%25""" + course_code + """%25%27+escape+%27%5C%27&orderby2=+kkxn%2Ckkxq%2Ckcdm%2Cjxbmc+asc+&seq2=&hiddenColumns=null&pager.offset=0&maxPageItemsSelect=20&maxPageItems2=20"""
        post_data = post_data.encode('utf-8')
        search_url = 'http://grdms.bit.edu.cn/yjs/yanyuan/py/pyjxjh.do?method=queryListing'
        headers = self.get_header_whole()
        r = requests.post(search_url, headers=headers, verify=False, data=post_data)
        pattern = "PKColumns\[0\]='(\S+)'"
        result = re.findall(pattern, r.text)
        if result:
            course_id = result[0]
            # 保存课程信息
            response = self.get_course_info_by_id(course_id)
            self.push_database(course_code, course_class_name, response)
            return {
                'status': True,
                'obj': response
            }
        else:
            return {
                'status': False,
                'msg': u'代码为%s、班级名称为%s的课程不存在' % (course_code, course_class_name)
            }

    def get_course_info_by_id(self, course_id):
        r"""
        根据课程id获取课程信息
        :param course_id:           课程id
        :return:                    课程信息，格式如下：
        {
            "id": 课程id,
            "code": 课程代码,
            "name": 课程名称,
            "type": 课程类别（如专业必修课）,
            "property": 课程性质（如专业必修课）,
            "hour": 学时,
            "credit": 学分
        }
        """
        detail_url = 'http://grdms.bit.edu.cn/yjs/yanyuan/py/pyjxjh.do?method=detail&id=' + course_id
        r = requests.get(detail_url, verify=False, headers=self.get_cookie_header())
        result = {
            "id": course_id
        }
        html = r.text
        # 代码和名称
        pattern1 = 'height="25">\s*(\S+)</td>'
        search_result = re.findall(pattern1, html)
        result['code'] = search_result[1]
        result['name'] = search_result[2]
        # 课程性质、学时、学分
        pattern2 = '<td>\s*(\S+)</td>'
        search_result = re.findall(pattern2, html)
        result['property'] = search_result[0]
        result['hour'] = search_result[1]
        result['credit'] = search_result[2]
        # 课程类别
        pattern3 = '<td\sheight="25">(\S+)</td>'
        search_result = re.findall(pattern3, html)
        result['type'] = search_result[0]
        return result

    def encode_chinese(self, course_class_name):
        r"""
        将中文转化成gbk编码的urlencode格式
        :param course_class_name:       班级名称
        :return:                        转化结果
        """
        course_class_name = str(course_class_name.encode('gbk')).replace('\\x', '%')
        course_class_name = course_class_name[2:][:-1]
        return course_class_name

    def get_course_info_online(self, course_code, course_class_name):
        post_url = 'http://grdms.bit.edu.cn/yjs/dwr/call/plaincall/YYPYCommonDWRController.getStdPyfaCourseList4Select.dwr'
        post_data = """callCount=1
page=/yjs/yanyuan/py/pyjxjh.do?method=stdSelectCourseEntry
httpSessionId=
scriptSessionId=
c0-scriptName=YYPYCommonDWRController
c0-methodName=getStdPyfaCourseList4Select
c0-id=0
c0-param0=string:%s
c0-param1=string:%s
c0-param2=string:%s
c0-e1=string:
c0-e2=string:%s
c0-e3=string:
c0-e4=string:
c0-e5=string:%s
c0-e6=string:%s
c0-e7=string:%s
c0-e8=string:%s
c0-e9=string:
c0-param3=Object_Object:{xh:reference:c0-e1, lb:reference:c0-e2, studentId:reference:c0-e3, id:reference:c0-e4, pycc:reference:c0-e5, xslb:reference:c0-e6, nj:reference:c0-e7, ldfs:reference:c0-e8, ptlx:reference:c0-e9}
c0-e10=string:%s
c0-e11=string:
c0-e12=string:
c0-e13=string:
c0-param4=Object_Object:{queryKcdm:reference:c0-e10, queryKcmc:reference:c0-e11, queryKclb:reference:c0-e12, queryKkdw:reference:c0-e13}
batchId=2
"""
        post_data = post_data % (
            self.basic_info['year'],
            quote(self.basic_info['term']),
            self.basic_info['pyfaId'],
            self.basic_info['lb'],
            quote(self.basic_info['education']),
            quote(self.basic_info['student_type']),
            self.basic_info['grade'],
            quote(self.basic_info['ldfs']),
            course_code
        )
        r = requests.post(post_url, data=post_data, headers=self.get_cookie_header())
        result = r.text.encode('utf-8').decode('unicode_escape')
        pattern = """s\d+\['TEACH_CLASS_ID'\]="(\S+)";s\d+.KKXN="(\S+)";s\d+.KKXQ="(\S+)";s\d+.KCDM="(\S+)";s\d+.JXBMC="(\S+)";s\d+.KCZWMC="(\S+)";s\d+.KCYWMC=(\S+);s\d+.KCXZ="(\S+)";s\d+.KCLB="(\S+)";s\d+.XS=(\S+);s\d+.XF=(\S+);s\d+.KSFS=(\S+);s\d+.KKDWMC="(\S+)";s\d+.RSXD=(\S+);s\d+.SFJX="(\S+)";s\d+.SJSKRS=(\S+);s\d+.SKJSXM="(\S+)";s\d+.SKSJDD="(.+)";s\d+.XKBZ=(\S+);s\d+.MXBZ=(\S+);"""
        result = re.findall(pattern, result)
        course = None
        for r in result:
            if r[4] == course_class_name:
                course = {
                    'id': r[0],
                    'year': r[1],
                    'term': r[2],
                    'code': r[3],
                    'class_name': r[4],
                    'name': r[5],
                    'property': r[7],
                    'type': r[8],
                    'hour': r[9],
                    'credit': r[10]
                }
                break
        if course is None:
            return {
                'status': False,
                'msg': u'课程代码为%s、教学班名称为%s的课程不存在' % (course_code, course_class_name)
            }
        return {
            'status': True,
            'obj': course
        }

    def add_course(self, course_code, course_class_name):
        r"""
        选课
        :param course_code:             课程代码
        :param course_class_name:       课程班名称
        :return:                        无
        """
        # 获取课程信息
        course_info = self.get_course_info(course_code, course_class_name)
        if not course_info['status']:
            print(course_info['msg'])
        else:
            course_info = course_info['obj']
            # 选课
            post_data = """callCount=1
page=/yjs/yanyuan/py/pyjxjh.do?method=stdSelectCourseEntry
httpSessionId=
scriptSessionId=
c0-scriptName=YYPYCommonDWRController
c0-methodName=pyJxjhSelectCourse
c0-id=0
c0-e1=string:%s
c0-e2=string:%s
c0-e3=string:%s
c0-e4=string:%s
c0-e5=string:%s
c0-e6=string:%s
c0-e7=string:%s
c0-e8=string:%s
c0-e9=string:
c0-param0=Object_Object:{xh:reference:c0-e1, lb:reference:c0-e2, studentId:reference:c0-e3, id:reference:c0-e4, pycc:reference:c0-e5, xslb:reference:c0-e6, nj:reference:c0-e7, ldfs:reference:c0-e8, ptlx:reference:c0-e9}
c0-e10=string:%s
c0-e11=string:%s
c0-e12=string:%s
c0-e13=string:%s
c0-e14=string:%s
c0-e15=string:%s
c0-e16=string:%s
c0-e17=string:%s
c0-e18=string:
c0-e19=string:
c0-e20=string:
c0-e21=string:
c0-e22=string:
c0-e23=string:
c0-param1=Object_Object:{studentId:reference:c0-e10, xh:reference:c0-e11, lb:reference:c0-e12, xn:reference:c0-e13, xq:reference:c0-e14, teachClassId:reference:c0-e15, kcdm:reference:c0-e16, kczwmc:reference:c0-e17, kcywmc:reference:c0-e18, kcxz:reference:c0-e19, kclb:reference:c0-e20, xs:reference:c0-e21, xf:reference:c0-e22, ksfs:reference:c0-e23}
batchId=12
"""
            post_data = post_data % (
                self.basic_info['stuid'],
                self.basic_info['lb'],
                self.basic_info['student_id'],
                self.basic_info['student_id'],
                quote(self.basic_info['education']),
                quote(self.basic_info['student_type']),
                self.basic_info['grade'],
                quote(self.basic_info['ldfs']),

                self.basic_info['student_id'],
                self.basic_info['stuid'],
                self.basic_info['lb'],
                self.basic_info['year'],
                quote(self.basic_info['term']),
                course_info['id'],
                course_info['code'],
                quote(course_info['name'])
            )
            add_url = 'http://grdms.bit.edu.cn/yjs/dwr/call/plaincall/YYPYCommonDWRController.pyJxjhSelectCourse.dwr'
            r = requests.post(add_url, data=post_data, verify=False, headers=self.get_cookie_header())
            html = r.text
            if html.find('success') != -1:
                return {
                    'status': True,
                    'msg': u'%s %s %s 选课成功！' % (course_code, course_class_name, course_info['name'])
                }
            else:
                try:
                    reason = html.split('"failure","')[1].split('"]);')[0]
                    reason = reason.encode('utf-8').decode('unicode_escape')
                    return {
                        'status': False,
                        'msg': u'%s %s %s 选课失败！: [%s]' % (course_code, course_class_name, course_info['name'], reason)
                    }
                except:
                    return {
                        'status':False,
                        'msg': u'接收到一次无效响应，重新开始...'
                    }

    def make_database(self):
        r"""
        提前创建选课数据库，这样选课的时候就可以少一步网络请求。
        :return:
        """
        if self.database != {}:
            print(u'课程数据库已存在')
            return
        print(u'正在提前创建选课数据库...')
        datas = course.read_courses()
        for data in datas:
            self.get_course_info(data['code'], data['class_name'])
            print(u'%s %s 完成' % (data['class_name'], data['code']))
        self.write_database()
        print(u'创建完成！')

    def add_course_loop(self, course_code, course_class_name, exit_on_error=True):
        r"""
        循环一直选课，直到选课完成
        :param course_code: 课程diamante
        :param course_class_name: 课程班名称
        :param exit_on_error:      遇到错误是否退出,默认是
        :return:
        """
        while True:
            result = self.add_course(course_code, course_class_name)
            if result is not None:
                print(result['msg'])
            else:
                print(u'选课失败了，可能是没开始选课。')
                time.sleep(0.5)
                continue
            if result['status']:
                # 选课成功之后当然要退出当前循环啦
                break
            else:
                if not exit_on_error:
                    continue
                # 如果是因为“时间冲突”和“已经选择”导致的失败，则不需要继续选课，记录下错误，并且退出循环即可
                key_words = [u'时间冲突', u'不能同时选择', u'已经选择']
                flag = False
                for word in key_words:
                    if result['msg'].find(word) != -1:
                        flag = True
                        break
                if flag:
                    # 记录当前错误
                    base_dir = os.path.dirname(__file__)
                    path = os.path.join(base_dir, "%s_%s_error.txt" % (course_class_name, course_code))
                    open(path, 'w+').write(result['msg'])
                    break

    def add_course_multi_threading(self, course_code, course_class_name):
        r"""
        创建多线程选课
        :param course_code:             课程代码
        :param course_class_name:       课程班名称
        :return:
        """
        _thread.start_new_thread(self.add_course_loop, (course_code, course_class_name, ))

    def add_all_course_using_threading(self, datas):
        r"""
        使用多线程选课，并且等待子线程退出
        :param datas:       所有课程
        :return:
        """
        threads = []
        for data in datas:
            threads.append(
                threading.Thread(target=self.add_course_loop, args=(
                    data['code'], data['class_name'], EXIT_ON_ERROR
                ))
            )
        for t in threads:
            t.setDaemon(True)
            t.start()
        for t in threads:
            t.join()

    def on_time(self, start_time, ahead=0):
        r"""
        判断当前时间是否超过选课时间
        :param start_time:      选课时间，字符串形式，格式为xxxx-xx-xx xx:xx:xx
        :param ahead:           提前的秒数，默认0不提前，为负值则延后
        :return:                是否可以开始选课
        """
        now = time.time()
        start_time = time.strptime(start_time, '%Y-%m-%d %H:%M:%S')
        start_time = time.mktime(start_time)
        return now >= start_time - ahead

    def enter_system(self):
        r"""
        进入选课系统，主要是需要获取用户之后，才能够选课。
        """
        base_dir = os.path.dirname(__file__)
        filename = os.path.join(base_dir, 'info_cache.pickle')
        if os.path.exists(filename):
            f = open(filename, 'rb')
            self.basic_info = pickle.load(f)
            print(u'使用缓存进入成功!')
            return self.basic_info
        else:
            i = 1
            while True:
                basic_info = course.get_basic_info()
                if basic_info['status']:
                    # 进入成功了
                    print(u'进入成功!')
                    f = open(filename, 'wb')
                    pickle.dump(basic_info['obj'], f)
                    return basic_info
                else:
                    print(u'进入不成功，重新尝试... %d' % i)
                    i = i + 1



if __name__ == '__main__':
    # -----------配置-------------------
    # 当选课返回异常（如已满、时间冲突）时，是否继续选这门课。
    # 如果是因为“时间冲突”和“已经选择”导致的失败，则不需要继续选课（说明不能选了，是自己的主观原因）。
    # 不影响已满的时候，已满继续抢。
    # 设置成False就行了。
    EXIT_ON_ERROR = False
    # ------------正文-----------------
    os.system(u'title 抢课系统')
    account = Course.read_account()
    course = Course(account['username'], account['password'])
    course.currentYear = account['currentYear']
    course.currentTerm = account['currentTerm']
    course.read_database()
    login_result = course.login(using_cache=True)
    if login_result['status']:
        os.system(u"title 【%s】 抢课系统" % course.stuid)
        print(u'【%s】 %s!' % (course.stuid, login_result['msg']))
        course.make_database()
        print(u'正在进入选课系统...')
        # ----------查看选课是否开始-------------- #
        basic_info = course.enter_system()
        # 多线程选课
        datas = course.read_courses()
        course.add_all_course_using_threading(datas)
    else:
        print(u'登录失败: ' + login_result['msg'])
    course.write_database()
    print(u'选课结束！')