#coding=utf-8
#author:Zhoubin

import os
import re
from multiprocessing import Pool,freeze_support
import pandas as pd

#get not PDF files as list
def get_file(dirpath):
    br = []
    x = os.listdir(dirpath)
    for fil in x:
        if os.path.splitext(fil)[1] != '.pdf':
            br.append(fil)
    return br

#write to xls
def openxls(model, rootdir, xlsworkbook):
    print 'appending >>> ', xlsworkbook
    r = re.compile(r'^\d')
    path = os.path.join(rootdir, xlsworkbook)
    xls = pd.ExcelFile(path)
    data = pd.read_excel(xls, sheetname=None, \
    header=None, index=None, encoding='gb18030', astype=str)
    for k in data.keys():
        if re.match(r, k):
            data[k].drop([0, 1, 2, 3], axis=0, inplace=True)
            data[k].dropna(axis=0, how='all', inplace=True)
            data[k].insert(0, 'motor_model', xlsworkbook[:-4])
            data[k][0].fillna(method='pad', inplace=True)
            data[k][1].fillna(method='pad', inplace=True)
            data[k].drop([2], axis=1, inplace=True)
            data[k][9] = data[k][9].str.replace(' ', '')
            data[k][9] = data[k][9].str.replace('\n', '')
            data[k].to_csv('d:\\{}.csv'.format(model), mode='a', header=None, index=None, encoding='gb18030')

if __name__ == '__main__':
    freeze_support()
    model = raw_input('Enter model series:>> ')
    rootdir = u'J:\\PIE Process Manual\\新工序文件(CMP)\\工序手冊\\{}\\'.format(model)
    ar = get_file(rootdir)
    p = Pool(4)
    for a in ar:
        p.apply(openxls, args=(model, rootdir, a, ))
    p.close()
    p.join()
    print 'D:\\{}.csv has been builded\n Copyright by Zhoubin'.format(model)
    os.system('pause.exe')

    '''
            d=data[k].iloc[4:-1,:]
            d.insert(0,'motor_model',xlsworkbook[:-4])
            d[0].fillna(method='pad',inplace=True)
            d[1].fillna(method='pad',inplace=True)
            d.drop([2],axis=1,inplace=True)
            d.to_csv('t.csv',mode='a',header=None,index=None,encoding='gb18030')

    for v in data.values():
        v.dropna(axis=0,how='all',inplace=True)
        v.insert(0,'motor_model',xlsworkbook[:-4])
        v[0].fillna(method='pad',inplace=True)
        v[1].fillna(method='pad',inplace=True)
        v.fillna(value='Null',axis=0,inplace=True)
        d=v[(v[0].str.startswith('A')) | (v[0].str.startswith('B')) | (v[0].str.startswith('C')) | (v[0].str.startswith('D')) | (v[0].str.startswith('E'))]
        d.to_csv('t.csv',mode='a',header=None,index=None,encoding='gb18030')
        #d=v[(v[1].str.contains(u'绕线',na=False)) | (v[1].str.contains(u'繞線',na=False))]
    '''