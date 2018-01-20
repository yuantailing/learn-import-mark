# -*- coding: utf-8 -*-

from __future__ import absolute_import
from __future__ import division
from __future__ import print_function
from __future__ import unicode_literals

import csv
import os
import re
import threading

from bs4 import BeautifulSoup
from six.moves import http_cookiejar as cookielib
from six.moves import queue
from six.moves import tkinter_font as font
from six.moves import tkinter_messagebox as messagebox
from six.moves import tkinter_tkfiledialog as tkfiledialog
from six.moves import tkinter_ttk as ttk
from six.moves import urllib
from six.moves.tkinter import *


class LearnFrame(ttk.Frame):
    def __init__(self, *args, **kwargs):
        ttk.Frame.__init__(self, *args, **kwargs)
        self.init_settings()

        self.cookie = cookielib.MozillaCookieJar(self.cookie_filename)
        if (os.path.isfile(self.cookie_filename)):
            self.cookie.load(ignore_discard=True, ignore_expires=True)
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cookie))

        self.students = []
        self.username = StringVar()
        self.password = StringVar()
        self.course_listvar = StringVar()
        self.assignment_listvar = StringVar()
        self.student_listvar = StringVar()

        ttk.Label(self, text='Username: ').grid(column=1, row=1, sticky=E)
        self.username_entry = ttk.Entry(self, width=10, textvariable=self.username)
        self.username_entry.grid(column=2, row=1, sticky=(W, E))
        ttk.Label(self, text='Password: ').grid(column=1, row=2, sticky=E)
        self.password_entry = ttk.Entry(self, textvariable=self.password, show='*')
        self.password_entry.grid(column=2, row=2, sticky=(W, E))
        self.login_button = ttk.Button(self, text='Login', command=self.login)
        self.login_button.grid(column=2, row=3, stick=E)

        self.course_listbox = Listbox(self, height=6, width=35, listvariable=self.course_listvar, relief='sunken')
        self.course_listbox.grid(column=3, row=1, rowspan=3, padx=(5, 0), stick=(N, W, E, S))
        self.course_listbox.configure(exportselection=False)
        self.course_listbox.bind('<<ListboxSelect>>', self.load_assignments)
        s = ttk.Scrollbar(self, orient=VERTICAL, command=self.course_listbox.yview)
        s.grid(column=4, row=1, rowspan=3, sticky=(N,S))
        self.course_listbox['yscrollcommand'] = s.set

        self.assignment_listbox = Listbox(self, height=6, width=35, listvariable=self.assignment_listvar, relief='sunken')
        self.assignment_listbox.grid(column=5, row=1, rowspan=3, padx=(5, 0), stick=(N, W, E, S))
        self.assignment_listbox.configure(exportselection=False)
        self.assignment_listbox.bind('<<ListboxSelect>>', self.load_students)
        s = ttk.Scrollbar(self, orient=VERTICAL, command=self.assignment_listbox.yview)
        s.grid(column=6, row=1, rowspan=3, sticky=(N,S))
        self.assignment_listbox['yscrollcommand'] = s.set

        self.student_listbox = Listbox(self, height=20, listvariable=self.student_listvar, relief='sunken', font=font.Font(family='Consolas', size=11))
        self.student_listbox.grid(column=1, row=4, columnspan=5, pady=(10, 5), stick=(N, W, E, S))
        self.student_listbox.configure(exportselection=False)
        s = ttk.Scrollbar(self, orient=VERTICAL, command=self.student_listbox.yview)
        s.grid(column=6, row=4, pady=(10, 5), sticky=(N,S))
        self.student_listbox['yscrollcommand'] = s.set

        self.readtxt_button = ttk.Button(self, text='读入成绩', command=self.read_txt)
        self.readtxt_button.grid(column=1, row=5, stick=E)
        self.dopublish_button = ttk.Button(self, text='确认发布', command=self.do_publish)
        self.dopublish_button.grid(column=2, row=5, stick=W)
        self.exportcsv_button = ttk.Button(self, text='导出成绩', command=self.export_csv)
        self.exportcsv_button.grid(column=3, row=5, stick=E)

        self.grid_columnconfigure(5, weight=1)
        self.grid_rowconfigure(4,weight=1)

        self.try_load_courses()

    def init_settings(self):
        self.dirname = '.'
        self.cookie_filename = 'cookie.txt'
        self.url_login = 'https://learn.tsinghua.edu.cn/MultiLanguage/lesson/teacher/loginteacher.jsp'
        self.url_logout = 'https://learn.tsinghua.edu.cn/MultiLanguage/lesson/logout.jsp'
        self.url_courselist = 'https://learn.tsinghua.edu.cn/MultiLanguage/lesson/student/MyCourse.jsp?typepage=3'
        self.url_assignmentlist = 'https://learn.tsinghua.edu.cn/MultiLanguage/lesson/teacher/hom_wk_brw.jsp?module_id=305&course_id={course_id}'
        self.url_studentlist = 'https://learn.tsinghua.edu.cn/MultiLanguage/lesson/teacher/hom_wk_reclist.jsp?id={assignment_id}&course_id={course_id}'
        self.url_homeworkmark = 'https://learn.tsinghua.edu.cn/MultiLanguage/lesson/teacher/hom_wk_recmark.jsp'
        self.url_homeworkdetail = 'https://learn.tsinghua.edu.cn/MultiLanguage/lesson/teacher/hom_wk_recdetail.jsp?rec_id={rec_id}&course_id={course_id}'

    def info(self, text, trace):
        messagebox.showinfo('Info', text)

    def error(self, text, trace):
        messagebox.showerror('Error', text)
        raise RuntimeError(trace)

    def html2unicode(self, text):
        return text.decode('utf-8', 'replace')

    def login(self):
        if self.login_button.cget('text') == 'Logout':
            response = self.opener.open(self.url_logout)
            self.cookie.save(ignore_discard=True, ignore_expires=False)
            self.try_load_courses()
        elif self.username.get() and self.password.get():
            request = urllib.request.Request(self.url_login)
            data = urllib.parse.urlencode({
                    'userid': self.username.get(),
                    'userpass': self.password.get(),
            })
            response = self.opener.open(request, data.encode('utf-8'))
            if 'window.alert' in self.html2unicode(response.read()):
                self.info('login failed', 'login failed')
            else:
                self.cookie.save(ignore_discard=True, ignore_expires=True)
                self.try_load_courses()

    def try_load_courses(self):
        response = self.opener.open(self.url_courselist)
        the_page = self.html2unicode(response.read())
        if not the_page.strip():
            self.username_entry.config(state=NORMAL)
            self.password_entry.config(state=NORMAL)
            self.login_button.config(text='Login')
            return False
        soup = BeautifulSoup(the_page, 'html.parser')
        self.course_list = []
        for a in soup.select('td table a'):
            match = re.search('course_id=(\d+)', a['href'])
            if not match:
                self.error('logic error', 'unexpected href')
            self.course_list.append({
                'id': int(match.group(1)),
                'name': a.text.strip(),
            })
        self.course_list.sort(key=lambda o: -o['id'])
        self.username_entry.config(state=DISABLED)
        self.password_entry.config(state=DISABLED)
        self.login_button.config(text='Logout')
        self.course_listvar.set(tuple([o['name'] for o in self.course_list]))
        self.course_listbox.selection_clear(0, END)
        self.assignment_listvar.set(())
        self.student_listvar.set(())
        self.students = []
        return True

    def load_assignments(self, *args):
        idxs = self.course_listbox.curselection()
        if 1 != len(idxs):
            self.error('logic error', 'current selection length != 1')
        self.course_id = self.course_list[idxs[0]]['id']
        self.course_name = self.course_list[idxs[0]]['name']
        response = self.opener.open(self.url_assignmentlist.format(course_id=self.course_id))
        the_page = self.html2unicode(response.read())
        soup = BeautifulSoup(the_page, 'html.parser')
        self.assignment_list = []
        for a in soup.select('a[href^="hom_wk_detail.jsp"]'):
            match = re.search('\?id=(\d+)', a['href'])
            if not match:
                self.error('logic error', 'unexpected href')
            self.assignment_list.append({
                'id': int(match.group(1)),
                'name': a.text.strip(),
            })
        self.assignment_listvar.set(tuple([o['name'] for o in self.assignment_list]))
        self.assignment_listbox.selection_clear(0, END)
        self.student_listvar.set(())
        self.students = []

    def update_student_listvar(self):
        def hanzi_pad(text, length):
            a = 0
            for c in text:
                if ord(c) < 256:
                    a += 1
                else:
                    a += 2
            return text + ' ' * max(0, length - a)
        l = []
        for o in self.students:
            if o['submitted']:
                s = '{:10s} {:s} (原成绩 {:s})'.format(o['id'], hanzi_pad(o['name'], 6), hanzi_pad(o['status_old'], 6))
                if 'score' in o:
                    if o['score'] is None:
                        s_score = 'None'
                    else:
                        s_score = '{:5.1f}'.format(o['score'])
                    s += '  导入 {:4s} {:5s}  {:s}'.format(o.get('published', ''), s_score, o['comment'])
            else:
                s = '{:s} {:s} 必须先交作业才能导入成绩'.format(o['id'], hanzi_pad(o['name'], 6))
            l.append(s)
        self.student_listvar.set(tuple(l))

    def load_students(self, *args):
        idxs = self.assignment_listbox.curselection()
        if 1 != len(idxs):
            self.error('logic error', 'current selection length != 1')
        self.assignment_id = self.assignment_list[idxs[0]]['id']
        self.assignment_name = self.assignment_list[idxs[0]]['name']
        response = self.opener.open(self.url_studentlist.format(assignment_id=self.assignment_id, course_id=self.course_id))
        the_page = self.html2unicode(response.read())
        soup = BeautifulSoup(the_page, 'html.parser')
        self.students = []
        for tr in soup.select('tr'):
            if tr.select('table'):
                continue
            if tr.select('a[href^="hom_wk_recdetail.jsp"]'):
                tds = tr.select('td')
                self.students.append({
                    'submitted': True,
                    'rec_id': int(re.search('rec_id=(\d+)', tds[2].select('a')[0]['href']).group(1)),
                    'id': tds[2].text.strip(),
                    'name': tds[3].text.strip(),
                    'status_old': tds[8].text.strip(),
                })
            elif tr.select('input[name="selectstu"]'):
                tds = tr.select('td')
                self.students.append({
                    'submitted': False,
                    'id': tds[2].text.strip(),
                    'name': tds[3].text.strip(),
                })
        self.update_student_listvar()
        self.student_listbox.selection_clear(0, END)

    def read_txt(self, *args):
        if not self.students:
            self.info('先选择作业，再读入成绩', 'read txt before select assignment')
            return
        f = tkfiledialog.askopenfile(mode='rb',
                                     initialdir=self.dirname,
                                     defaultextension='*.txt',
                                     filetypes=(('Text file', '*.txt'), ('All types', '*.*')))
        if f is None:
            return
        with f:
            self.dirname = os.path.dirname(f.name)
            content = f.read()
        try:
            content = content.decode('utf-8')
        except:
            try:
                content = content.decode('cp936')
            except:
                self.info('文件编码错误，请使用 utf-8 或 cp936 编码（优先 UTF-8）', 'input file encoding error')
                return
        imported = []
        for line in content.strip().splitlines():
            line = line.strip()
            if line == '':
                continue
            a = line.split(None, 2)
            if len(a) < 2:
                self.info('格式错误，每行应为“学号 分数 评语”，分数为空可以用半角减号“-”代替，评语可省略,分数可保留 1 位小数', 'ill formated')
                return
            id = a[0].strip()
            comment = a[2].strip() if 2 < len(a) else ''
            try:
                b = a[1].strip()
                if b == '-':
                    score = None
                else:
                    score = float(b)
            except:
                self.info('格式错误，每行应为“学号 分数 评语”，分数为空可以用半角减号“-”代替，评语可省略,分数可保留 1 位小数', 'ill formated')
                return
            if score is not None and (score < 0 or score > 100):
                self.info('成绩必须在 0-100 范围内', 'ill formated')
                return
            imported.append([id, score, comment])

        for o in self.students:
            for k in ('score', 'comment', 'published'):
                if k in o:
                    o.pop(k)
        for elem in imported:
            for o in self.students:
                if o['id'] == elem[0]:
                    o['score'] = elem[1]
                    o['comment'] = elem[2]
        self.update_student_listvar()

    def do_publish(self, *args):
        num_update = 0
        for o in self.students:
            if o['submitted'] and 'score' in o:
                num_update += 1
        if 0 == num_update:
            self.info('先读入成绩，再确认发布', 'read txt before select assignment')
            return
        if not messagebox.askyesno('Publish', '更新 【{:s}】 【{:s}】 的 【{:d}】 项分数。\n该操作不可撤销。您已经检查好了吗？'.format(
                    self.course_name, self.assignment_name, num_update)):
            return
        to_lock = [
            self.login_button,
            self.course_listbox,
            self.assignment_listbox,
            self.readtxt_button,
            self.dopublish_button,
            self.exportcsv_button,
        ]
        def target(course_id):
            for o in self.students:
                if 'published' in o:
                    o['published'] = False
            self.update_student_listvar()
            for o in self.students:
                if o['submitted'] and 'score' in o:
                    try:
                        response = self.opener.open(self.url_homeworkmark, urllib.parse.urlencode({
                            'post_rec_mark': '{:f}'.format(o['score']) if o['score'] is not None else '',
                            'post_rec_reply_detail': o['comment'],
                            'course_id': course_id,
                            'post_homewkrec_id': o['rec_id']
                        }).encode())
                        assert 200 == response.code
                        o['published'] = '[OK]'
                    except:
                        o['published'] = 'FAIL'
            q.put(None)
        for elem in to_lock:
            elem.config(state=DISABLED)
        q = queue.Queue()
        def update(q):
            try:
                msg = q.get_nowait()
                self.update_student_listvar()
                for elem in to_lock:
                    elem.config(state=NORMAL)
            except queue.Empty:
                self.after(100, update, q)
        t = threading.Thread(target=target, args=(self.course_id,))
        t.start()
        self.after(100, update, q)

    def export_csv(self, *args):
        to_lock = [
            self.login_button,
            self.course_listbox,
            self.assignment_listbox,
            self.readtxt_button,
            self.dopublish_button,
            self.exportcsv_button,
        ]
        def target(course_id):
            rows = []
            all_right = True
            for o in self.students:
                if o['submitted']:
                    try:
                        response = self.opener.open(self.url_homeworkdetail.format(
                            rec_id=o['rec_id'], course_id=course_id))
                        assert 200 == response.code
                        the_page = self.html2unicode(response.read())
                        soup = BeautifulSoup(the_page, 'html.parser')
                        student_id = soup.select('#post_rec_student_id')[0]['value'].strip()
                        assert student_id == o['id']
                        answer = soup.select('textarea[readonly]')[0].text
                        mark = soup.select('#post_rec_mark')[0]['value'].strip()
                        reply = soup.select('#post_rec_reply_detail')[0].text
                        rows.append([
                            o['id'],
                            o['name'],
                            answer.strip(),
                            mark,
                            reply.strip(),
                        ])
                    except:
                        rows.append([
                            o['id'],
                            o['name'],
                        ])
                        all_right = False
            if not all_right:
                q.put(('读取详细评语失败', 'open homework rec failed'))
                return
            with f:
                writer = csv.writer(f, dialect='excel')
                writer.writerow([
                    'id',
                    'name',
                    'answer',
                    'mark',
                    'reply',
                ])
                writer.writerows(rows)
            q.put(None)

        try:
            fn = tkfiledialog.asksaveasfilename(initialdir=self.dirname,
                                                defaultextension='*.csv',
                                                filetypes=(('CSV file', '*.csv'), ('All types', '*.*')))
            if fn is None:
                return
            f = open(fn, 'w', encoding='utf-8-sig', newline='')
        except IOError as e:
            self.info(str(e), 'IOError')
            return
        self.dirname = os.path.dirname(f.name)
        for elem in to_lock:
            elem.config(state=DISABLED)
        q = queue.Queue()
        def update(q):
            try:
                msg = q.get_nowait()
                if msg is not None:
                    self.info(*msg)
                for elem in to_lock:
                    elem.config(state=NORMAL)
            except queue.Empty:
                self.after(100, update, q)
        t = threading.Thread(target=target, args=(self.course_id,))
        t.start()
        self.after(100, update, q)

root = Tk()
root.title('Import')

mainframe = LearnFrame(root, padding=(10, 5, 10, 5))
mainframe.grid(column=0, row=0, sticky=(N,W,E,S))
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0,weight=1)

root.mainloop()
