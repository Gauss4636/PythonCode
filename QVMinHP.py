# from django.shortcuts import render, HttpResponse
# import json
# from django.http import JsonResponse
# from html.parser import HTMLParser
# import bs4
import requests, re
from bs4 import BeautifulSoup
from requests.auth import HTTPDigestAuth
import xlwt


# 获取信息写入列表
def Addinfotolist(info_name, file_name):
    info = file_name.findAll(info_name)
    list = []
    for i in info:
        list.append(i.string)
    return list


# 将存储信息的列表拼接成excel直接使用的列表形式
def ExcellistCreate(infolist, newlist):
    for i in range(0, len(infolist)):
        str1 = ''
        for j in range(0, len(infolist[i])):
            if j < len(infolist[i]) - 1:
                str1 += infolist[i][j] + ','
            else:
                str1 += infolist[i][j]
        newlist.append(str1)
    return newlist


# 通过主机池id查询所有主机
def Qhostinfo():
    url = "http://10.10.0.7:8080/cas/casrs/hostpool/host/1?offset=0&limit=200"
    resp = requests.get(url, auth=HTTPDigestAuth('yd_gaosi', 'jsdk@123'))
    hostinfohtml = resp.text
    hostinfohtml = re.sub("name", "namete", hostinfohtml, count=0, flags=0)
    hostinfohtml = re.sub("vmNum", "vmnum", hostinfohtml, count=0, flags=0)
    hostinfohtml = re.sub("vmRunCount", "vmruncount", hostinfohtml, count=0, flags=0)
    hostinfohtml = re.sub("vmShutoff", "vmshutoff", hostinfohtml, count=0, flags=0)
    hostinfohtml = re.sub("cpuRate", "cpurate", hostinfohtml, count=0, flags=0)
    hostinfohtml = re.sub("memRate", "memrate", hostinfohtml, count=0, flags=0)
    soup_string = BeautifulSoup(hostinfohtml, "html.parser")
    # 主机id
    host_id = Addinfotolist('id', soup_string)
    # 主机名
    host_name = Addinfotolist('namete', soup_string)
    # 主机管理ip
    host_ip = Addinfotolist('ip', soup_string)
    # 主机上虚拟机数量
    host_vmnum = Addinfotolist('vmnum', soup_string)
    # 主机上正在运行虚拟机数量
    host_vmruncount = Addinfotolist('vmruncount', soup_string)
    # 主机上关闭的虚拟机数量
    host_vmshutoff = Addinfotolist('vmshutoff', soup_string)
    # 主机的cpu利用率(保留两位小数的特殊处理)
    host_cpurate_info = soup_string.findAll('cpurate')
    host_cpurate = []
    for i in host_cpurate_info:
        host_cpurate.append((round(float(i.string), 2)))
    # 主机的内存利用率(保留两位小数的特殊处理)
    host_memrate_info = soup_string.findAll('memrate')
    host_memrate = []
    for i in host_memrate_info:
        host_memrate.append(round(float(i.string), 2))
    # 主机的平台
    host_version = Addinfotolist('version', soup_string)
    # 主机的状态
    host_status = Addinfotolist('status', soup_string)
    # 主机磁盘大小（MB转换为GB）
    host_storage = []
    host_storage_info = soup_string.findAll('storage')
    for i in host_storage_info:
        host_storage.append(round(int(i.string) / 1024, 2))

    hostinfo_dic = {}
    hostdic_list = []
    for i in range(0, len(host_id)):
        hostinfo_dic[host_id[i]] = (
            host_name[i], host_ip[i], host_vmnum[i], host_vmruncount[i], host_vmshutoff[i], host_cpurate[i],
            host_memrate[i], host_status[i], host_storage[i])

    # 生成Ecel表格的数据列表
    hostall_cpu = []
    hostall_cpumodel = []
    hostall_memory = []
    hostall_ethinfo_key = []
    hostall_ethinfo_value = []
    hostall_storageinfo_key = []
    hostall_storageinfo_value = []
    hostall_mac = []
    hostall_flowtime = []
    hostall_receiveflow = []
    hostall_sendflow = []
    hostall_stex_name = []
    hostall_stex_rdst = []
    hostall_stex_wrst = []
    hostall_stex_rdreq = []
    hostall_stex_wrreq = []
    exst_receiveflow = []
    exmg_receiveflow = []
    exapp_receiveflow = []
    exst_sendflow = []
    exmg_sendflow = []
    exapp_sendflow = []
    host_ethnumber = []

    # 通过主机ID查询单个主机详细信息
    for id in host_id:
        singlehost_dic = {}
        # 根据每个主机的信息字典生成一个每次循环显示单一主机信息的列表
        host_list = list(hostinfo_dic[id])
        singlehost_dic["hostip"] = host_list[1]
        singlehost_dic["vmnumber"] = host_list[2]
        singlehost_dic["vmrunnum"] = host_list[3]
        singlehost_dic["vmshutoffnum"] = host_list[4]
        singlehost_dic["cpurate"] = host_list[5]
        singlehost_dic["memrate"] = host_list[6]
        singlehost_dic["status"] = host_list[7]
        singlehost_dic["storagesize"] = host_list[8]
        # 通过主机id查询主机详细信息
        host_url1 = "http://10.10.0.7:8080/cas/casrs/host/id/%s" % id
        resp1 = requests.get(host_url1, auth=HTTPDigestAuth('yd_gaosi', 'jsdk@123'))
        hostinfohtml_detail = resp1.text
        hostinfohtml_detail = re.sub("name", "ethname", hostinfohtml_detail, count=0, flags=0)
        hostinfohtml_detail = re.sub("cpuCount", "cpucount", hostinfohtml_detail, count=0, flags=0)
        hostinfohtml_detail = re.sub("cpuModel", "cpumodel", hostinfohtml_detail, count=0, flags=0)
        hostinfohtml_detail = re.sub("memorySize", "memorysize", hostinfohtml_detail, count=0, flags=0)
        hostinfohtml_detail = re.sub("pNIC", "pnic", hostinfohtml_detail, count=0, flags=0)
        hostinfohtml_detail = re.sub("macAddr", "macaddr", hostinfohtml_detail, count=0, flags=0)
        soup_string_detail = BeautifulSoup(hostinfohtml_detail, "html.parser")
        # 主机cpu
        host_cpucount = Addinfotolist('cpucount', soup_string_detail)
        singlehost_dic['cpu'] = host_cpucount
        # 主机cpu型号
        host_cpumodel = Addinfotolist('cpumodel', soup_string_detail)
        singlehost_dic['cpumodel'] = host_cpumodel
        # 主机内存
        host_memorysize = []
        host_memsizeinfo = soup_string_detail.findAll('memorysize')
        for i in host_memsizeinfo:
            host_memorysize.append(str(round(int(i.string) / 1024, 2)))
        singlehost_dic['memorysize'] = host_memorysize
        # 获取网卡名称（多个网卡）
        soup_eth_name = soup_string_detail.findAll("pnic")
        host_eth_name = []
        for i in soup_eth_name:
            host_eth_name.append(i.ethname.string)
        # 主机网卡mac地址
        host_macaddr = []
        for i in soup_eth_name:
            host_macaddr.append(i.macaddr.string)
        ethinfo = dict(zip(host_eth_name, host_macaddr))
        singlehost_dic['ethinfo'] = ethinfo
        # 主机存储信息
        host_url1 = "http://10.10.0.7:8080/cas/casrs/host/id/%s/monitor" % id
        resphostmonitor = requests.get(host_url1, auth=HTTPDigestAuth('yd_gaosi', 'jsdk@123'))
        hostinfohtml_storage = resphostmonitor.text
        soup_string_hoststorage = BeautifulSoup(hostinfohtml_storage, "html.parser")
        # 主机磁盘名称
        host_device = Addinfotolist('device', soup_string_hoststorage)
        # 主机磁盘读速率
        host_rd_stat = Addinfotolist('rd_stat', soup_string_hoststorage)
        # 主机磁盘写速率
        host_wr_stat = Addinfotolist('wr_stat', soup_string_hoststorage)
        # 主机磁盘读IOPS
        host_rd_req = Addinfotolist('rd_req', soup_string_hoststorage)
        # 主机磁盘写IOPS
        host_wr_req = Addinfotolist('wr_req', soup_string_hoststorage)

        # 生成主机存储IO数据字典
        hoststorage = {}
        for i in range(0, len(host_device)):
            hoststorage[host_device[i]] = (host_rd_stat[i], host_wr_stat[i], host_rd_req[i], host_wr_req[i])
        singlehost_dic['storageinfo'] = hoststorage

        # Excel使用的分开形式的磁盘读写信息信息列表
        hostall_stex_name.append(host_device)
        hostall_stex_rdst.append(host_rd_stat)
        hostall_stex_wrst.append(host_wr_stat)
        hostall_stex_rdreq.append(host_rd_req)
        hostall_stex_wrreq.append(host_wr_req)

        # 通过mac地址查询网卡的流量
        eth_trafic = {}
        hostsingle_flowtime = []
        hostsingle_receiveflow = []
        hostsingle_sendflow = []
        for mac in host_macaddr:
            mac_url = "http://10.10.0.7:8080/cas/casrs/host/pnic/traffic?mac=%s" % mac
            resphostmac = requests.get(mac_url, auth=HTTPDigestAuth('yd_gaosi', 'jsdk@123'))
            hostinfohtml_mac = resphostmac.text
            hostinfohtml_mac = re.sub("trendRate", "trendrate", hostinfohtml_mac, count=0, flags=0)
            soup_string_hostmac = BeautifulSoup(hostinfohtml_mac, "html.parser")
            trendrate = soup_string_hostmac.findAll('rates')
            receive_flow_dic = {}
            send_flow_dic = {}
            hostflow_time = ''
            host_receiveflow_rate = ''
            host_sendflow_rate = ''
            num = 0
            number = 0
            # time参数发送流量和接收流量均相同，个数都为20个。共用同一个时间列表
            for i in trendrate:
                num += 1
                if num == 20:
                    hostflow_time = i.time.string
            # 查找取得所有的流量，前20个数据为接收流量后20个为发送流量
            for i in trendrate:
                number += 1
                if number == 20:
                    host_receiveflow_rate = i.rate.string
                elif number == 40:
                    host_sendflow_rate = i.rate.string
            # 拼接流量和时间的字典后加上对应的mac地址生成新的字典
            receive_flow = hostflow_time + ':' + host_receiveflow_rate
            send_flow = hostflow_time + ':' + host_sendflow_rate
            receive_flow_dic["receiveflow"] = receive_flow
            send_flow_dic["sendflow"] = send_flow
            eth_trafic[mac] = [receive_flow_dic, send_flow_dic]

            # 生成excel用网卡流量用时间列表和流量列表
            hostsingle_flowtime.append(hostflow_time)
            hostsingle_receiveflow.append(host_receiveflow_rate)
            hostsingle_sendflow.append(host_sendflow_rate)

        # 将单一主机的所有网卡的流量字典组成的列表传给单一主机信息汇总的singlehost_dic
        singlehost_dic['flow'] = eth_trafic
        hostdic_list.append(singlehost_dic)

        # 加入每个主机的相应信息到Excel使用的表格
        hostall_cpu.append(host_cpucount)
        hostall_cpumodel.append(host_cpumodel)
        hostall_memory.append(host_memorysize)
        hostall_ethinfo_key.append(list(ethinfo.keys()))
        hostall_ethinfo_value.append(list(ethinfo.values()))
        hostall_storageinfo_key.append(list(hoststorage.keys()))
        hostall_storageinfo_value.append(list(hoststorage.values()))
        hostall_mac.append(host_macaddr)
        hostall_flowtime.append(hostsingle_flowtime)
        hostall_receiveflow.append(hostsingle_receiveflow)
        hostall_sendflow.append(hostsingle_sendflow)

    # 根据不同网卡聚合成三种网络：存储、管理、业务 根据网卡的顺序来判断相应的流量排列顺序
    for ethname in hostall_ethinfo_key:
        host_flownumber = {}
        for j in range(0, len(ethname)):
            if ethname[j] == 'eth0':
                host_flownumber['eth0'] = j
            elif ethname[j] == 'eth1':
                host_flownumber['eth1'] = j
            elif ethname[j] == 'eth2':
                host_flownumber['eth2'] = j
            elif ethname[j] == 'eth3':
                host_flownumber['eth3'] = j
            elif ethname[j] == 'eth4':
                host_flownumber['eth4'] = j
            elif ethname[j] == 'eth5':
                host_flownumber['eth5'] = j
        host_ethnumber.append(host_flownumber)

    # 生成三种网络的接收流量数据0、1网卡为存储网络，2、3为管理网络，4、5为业务网络
    for i in hostall_receiveflow:
        j = 0
        st_reflow = round(float(i[host_ethnumber[j]['eth0']]) + float(i[host_ethnumber[j]['eth1']]), 2)
        mg_reflow = round(float(i[host_ethnumber[j]['eth2']]) + float(i[host_ethnumber[j]['eth3']]), 2)
        ap_reflow = round(float(i[host_ethnumber[j]['eth4']]) + float(i[host_ethnumber[j]['eth5']]), 2)
        exst_receiveflow.append(st_reflow)
        exmg_receiveflow.append(mg_reflow)
        exapp_receiveflow.append(ap_reflow)
        j += 1

    # 生成三种网络的发送流量数据
    for i in hostall_sendflow:
        j = 0
        st_sdflow = round(float(i[host_ethnumber[j]['eth0']]) + float(i[host_ethnumber[j]['eth1']]), 2)
        mg_sdflow = round(float(i[host_ethnumber[j]['eth2']]) + float(i[host_ethnumber[j]['eth3']]), 2)
        ap_sdflow = round(float(i[host_ethnumber[j]['eth4']]) + float(i[host_ethnumber[j]['eth5']]), 2)
        exst_sendflow.append(st_sdflow)
        exmg_sendflow.append(mg_sdflow)
        exapp_sendflow.append(ap_sdflow)
        j += 1

    # 生成excel用网卡名称和mac地址对应的列表并以主机为单位写入同一单元格
    hostall_ethinfo = []
    y = 0
    for key in hostall_ethinfo_key:
        x = 0
        str1 = ''
        for value in hostall_ethinfo_value[y]:
            l1 = key[x] + "->" + value
            if x < len(key) - 1:
                str1 += l1 + ','
            else:
                str1 += l1
            x += 1
        y += 1
        hostall_ethinfo.append(str1)

    # 生成excel用磁盘信息将磁盘名称和磁盘读写速率生成对应列表一主机为单位写入一个单元格
    hostall_storageinfo = []
    y = 0
    for key in hostall_storageinfo_key:
        x = 0
        str1 = ''
        for value in hostall_storageinfo_value[y]:
            l1 = '磁盘读速率' + ':' + value[0]
            l2 = '磁盘写速率' + ':' + value[1]
            l3 = '磁盘读IOPS' + ':' + value[2]
            l4 = '磁盘写IOPS' + ':' + value[3]
            if x < len(key) - 1:
                str1 += key[x] + ':' + l1 + ',' + l2 + ',' + l3 + ',' + l4 + ','
            else:
                str1 += key[x] + ':' + l1 + ',' + l2 + ',' + l3 + ',' + l4
            x += 1
        y += 1
        hostall_storageinfo.append(str1)

    hostall_reflowinfo = []
    hostall_seflowinfo = []
    eth_time = hostall_flowtime[0][0]
    y = 0
    for mac in hostall_mac:
        x = 0
        str1 = ''
        for flow in hostall_receiveflow[y]:
            if x == len(hostall_receiveflow[y]) - 1:
                str1 += mac[x] + '->' + eth_time + ':' + flow
            else:
                str1 += mac[x] + '->' + eth_time + ':' + flow + ','
            x += 1
        y += 1
        hostall_reflowinfo.append(str1)

    # 生成excel用网卡流量中最近的一次时间采集点的接收和发送流量情况
    y = 0
    for mac in hostall_mac:
        x = 0
        str1 = ''
        for flow in hostall_sendflow[y]:
            if x == len(hostall_sendflow[y]) - 1:
                str1 += mac[x] + '->' + eth_time + ':' + flow
            else:
                str1 += mac[x] + '->' + eth_time + ':' + flow + ','
            x += 1
        y += 1
        hostall_seflowinfo.append(str1)

    # 所有单一主机信息字典拼接主机名生成全部信息字典
    hostallinfo = dict(zip(host_name, hostdic_list))

    # 调用函数生成主机存储iops速率独立的单元格形式的列表
    host_stname = []
    host_strdst = []
    host_stwrst = []
    host_strdreq = []
    host_stwrreq = []
    ExcellistCreate(hostall_stex_name, host_stname)
    ExcellistCreate(hostall_stex_rdst, host_strdst)
    ExcellistCreate(hostall_stex_wrst, host_stwrst)
    ExcellistCreate(hostall_stex_rdreq, host_strdreq)
    ExcellistCreate(hostall_stex_wrreq, host_stwrreq)

    # 生成Excel主机信息表格
    work_book = xlwt.Workbook()
    sheet = work_book.add_sheet('HostInfoList')
    # 表头生成
    titlelist = ['主机名称', '管理IP地址', '虚拟机数量', '运行虚拟机数量', '停止虚拟机数量', 'CPU', 'CPU型号', '内存大小(GB)', 'CPU占用率(%)', '内存占用率(%)',
                 '主机状态', '磁盘大小(GB)', '磁盘名称', '磁盘读速率', '磁盘写速率', '磁盘读IOPS', '磁盘写IOPS', '磁盘信息汇总', '网卡信息', '接收流量汇总',
                 '发送流量汇总']
    titlelist2 = ['存储网络', '管理网络', '业务网络', '存储网络', '管理网络', '业务网络']
    for i in range(0, len(titlelist)):
        sheet.write_merge(0, 1, i, i, titlelist[i])
    for i in range(0, len(titlelist2)):
        sheet.write(1, len(titlelist) + i, titlelist2[i])
    sheet.write_merge(0, 0, 21, 23, '网卡接收数据(MB)')
    sheet.write_merge(0, 0, 24, 26, '网卡发送数据(MB)')

    # 将数据注入Excel表格中
    for i in range(0, len(host_id)):
        datalist = [host_name, host_ip, host_vmnum, host_vmruncount, host_vmshutoff, hostall_cpu, hostall_cpumodel,
                    hostall_memory, host_cpurate, host_memrate, host_status, host_storage,
                    host_stname, host_strdst, host_stwrst, host_strdreq, host_stwrreq, hostall_storageinfo,
                    hostall_ethinfo, hostall_reflowinfo, hostall_seflowinfo, exst_receiveflow, exmg_receiveflow,
                    exapp_receiveflow, exst_sendflow, exmg_sendflow, exapp_sendflow]
        k = 0
        for j in datalist:
            sheet.write(i + 2, k, j[i])
            k += 1
    #保存Excel文件到指定目录
    work_book.save('D:\PythonWork\主机信息.xlsx')
    print('主机信息.xlsx 已生成！')

    return hostallinfo

Qhostinfo()

