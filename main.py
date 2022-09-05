from lxml import etree
import re
import time
import requests
import openpyxl
import threading

headers = {
    'User-Agent':
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'
}
url = 'https://jinshuju.net/****'

year_find = re.compile(
    r'<tr data-field="field_1" data-value="(.*?)" data-api-code="field_1" data-field-type="text-field">')
title_find = re.compile(
    r'<tr data-field="field_2" data-value="(.*?)" data-api-code="field_2" data-field-type="text-field">')
name_find = re.compile(
    r'<tr data-field="field_3" data-value="(.*?)" data-api-code="field_3" data-field-type="text-field">')
id_find = re.compile(
    r'<tr data-field="field_4" data-value="(.*?)" data-api-code="field_4" data-field-type="text-field">')

data = []
student_id = []
range_list = [[0, 100], [100, 200], [200, 300], [300, 400]]
threads_list = []


def get_html(id):
    origin_html = requests.get(url=str(url + id), headers=headers)
    return origin_html


def get_ids():
    wb = openpyxl.load_workbook('****.xlsx').active
    column = wb['A2':'B5491']
    for (c1, c2) in column:
        student_id.append([c1.value, c2.value])


def work(thread_id):
    start = thread_id * 2745
    finish = thread_id * 2745 + 2745
    for i in student_id[start:finish]:
        person = []
        html = get_html(str(i[0]))
        person.append(str(i[0]))
        person.append(str(i[1]))
        person += year_find.findall(html.text)
        person += title_find.findall(html.text)

        if len(person) == 2:
            html_tree = etree.HTML(html.text)
            try:
                person.append(html_tree.xpath('//*[@id="entries_table"]/tbody/tr[1]/td[1]/div/@title')[0])
                person.append(html_tree.xpath('//*[@id="entries_table"]/tbody/tr[1]/td[2]/div/@title')[0])
                person.append(html_tree.xpath('//*[@id="entries_table"]/tbody/tr[2]/td[1]/div/@title')[0])
                person.append(html_tree.xpath('//*[@id="entries_table"]/tbody/tr[2]/td[2]/div/@title')[0])
            except:
                pass
        data.append(person)


if __name__ == '__main__':
    start_time = time.time()

    get_ids()
    print('finish read excel')

    for i in range(2):
        thread = threading.Thread(target=work, args=(i,))
        thread.start()
        threads_list.append(thread)
    for i in threads_list:
        i.join()

    workbook = openpyxl.Workbook()
    wb_active = workbook.active
    for i in data:
        wb_active.append(i)
    workbook.save('multiple.xlsx')

    finish_time = time.time()
    print('run time:', finish_time - start_time)
