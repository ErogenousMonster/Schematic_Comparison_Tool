# coding=utf-8
import copy
import os
import re
import xlsxwriter
import pandas as pd
from pathlib import Path
import xlwings as xw
from math import ceil


# Class part


# 读取 pstxnet.dat，pstxprt.dat，pstchip.dat 及 EXP文件的模块
class ExtractIOData:
    """对DSN导出的报告进行数据处理"""

    def __init__(self):
        root_path = os.getcwd()
        self._root_path = os.path.join(root_path, 'input')
        self._output_path = os.path.join(root_path, 'output')
        self._output_excel_path = os.path.join(self._output_path, 'Initial_GPIO_Table.xlsx')

    def extract_pstxnet(self, dsn_path):
        """提取pstxnet.dat的数据"""
        all_net_list_ = []
        net_component_dict_ = {}
        net_component_list_ = []

        try:
            with open(os.path.join(dsn_path, 'pstxnet.dat'), 'r') as file1:
                content1 = file1.read().split('NET_NAME')
                for ind1 in range(len(content1)):
                    content1[ind1] = content1[ind1].split('\n')
                for x in content1:
                    component_list = []
                    all_net_list_.append(x[1][1:-1])
                    for y_idx in range(len(x)):
                        if x[y_idx].find('NODE_NAME') > -1:
                            component_list.append([x[y_idx].split('NODE_NAME\t')[-1].
                                                  split(' ')[0], x[y_idx + 2].split("'")[1]])
                    component_flatten_list = list(flatten([[x[1][1:-1]] + component_list]))
                    # print('component_flatten_list', component_flatten_list)
                    net_component_dict_[x[1][1:-1]] = component_flatten_list
                    net_component_list_.append(component_flatten_list)
                all_net_list_ = all_net_list_[1:]
            # print(self.net_component_list_)
            return all_net_list_, net_component_list_, net_component_dict_
        except FileNotFoundError:
            error_message = 'Missing pstxnet.dat file'
            create_error_message(self._output_excel_path, error_message)
            raise FileNotFoundError

    def extract_pstxprt(self, dsn_path):
        """提取pstxprt.dat的数据"""
        all_component_list_ = []
        component_page_dict_ = {}
        primitive_name_component_name_dict_ = {}
        component_name_primitive_name_dict_ = {}
        try:
            with open(os.path.join(dsn_path, 'pstxprt.dat'), 'r') as file2:
                content2 = file2.read().split('PART_NAME')

                primitive_list = []

                for ind2 in range(len(content2)):
                    content2[ind2] = content2[ind2].split('\n')

                for x in content2:
                    # print(x)
                    # print('\n')
                    component = x[1].split(' ')[1]
                    all_component_list_.append(component)
                    if x[0] == '':
                        pattern2 = re.compile(r".*?'(.*?)'.*?")
                        # pattern3 = re.compile(r".*?@.*?@(.*?)\..*?")
                        component1 = x[1].split(' ')[1]
                        component_name_primitive_name_dict_[component1] = pattern2.findall(x[1])[0]
                        if pattern2.findall(x[1])[0] not in primitive_name_component_name_dict_.keys():
                            primitive_name_component_name_dict_[pattern2.findall(x[1])[0]] = component1

                        # if pattern3.findall(x[5])[0].upper() == 'RESISTOR':
                        #     self.all_res_list_.append(component1)

                    # pattern = re.compile(r".*?_.*?_(.*?)_.*?")
                    # res_val = pattern.findall(x[1].split(' ')[2])

                    # if res_val:
                    #     self.all_res_dict_[component] = res_val[0]

                    if x[0] == '':
                        primitive_list.append(x[1].split("\'")[1])
                    if 'page' in x[6]:
                        page_now = x[6].split(':')[-1].split('_')[0]
                    else:
                        page_now = x[7].split(':')[-1].split('_')[0]

                    component_page_dict_[all_component_list_[-1]] = page_now[4:]
                    # if self.component_page_dict_.get(page_now):
                    #     self.component_page_dict_[page_now] += [self.all_component_list_[-1]]
                    # else:
                    #     self.component_page_dict_[page_now] = [self.all_component_list_[-1]]

            return all_component_list_, component_page_dict_, primitive_name_component_name_dict_, \
                   component_name_primitive_name_dict_
        except FileNotFoundError:
            error_message = 'Missing pstxprt.dat file'
            create_error_message(self._output_excel_path, error_message)
            raise FileNotFoundError

    def extract_pstchip(self, dsn_path):
        primitive_name_pin_num_dict_ = {}
        primitive_name_pin_list_dict = {}
        try:
            with open(os.path.join(dsn_path, 'pstchip.dat'), 'r') as file5:
                # 获取ic及其pin的数量信息
                content = file5.read().split('end_primitive')
                pattern1 = re.compile(r".*?primitive '(.*?)'")
                pattern2 = re.compile(r".*?    '(.*?)':")

                for c_item in content:
                    key_item = pattern1.findall(c_item)
                    pin_list = pattern2.findall(c_item)
                    if key_item:
                        primitive_name_pin_num_dict_[key_item[0]] = c_item.count('PIN_NUMBER')
                        primitive_name_pin_list_dict[key_item[0]] = pin_list

                return primitive_name_pin_num_dict_, primitive_name_pin_list_dict

        except FileNotFoundError:
            error_message = 'Missing pstchip.dat file'
            create_error_message(self._output_excel_path, error_message)
            raise FileNotFoundError

    def extract_exp(self, dsn_path, foxc_flag=True):
        component_mfg_dict = {}
        # 寻找.exp文件
        file_name = ''
        for x in os.listdir(dsn_path):
            (shortname, extension) = os.path.splitext(x)
            # print(extension)
            if extension == '.EXP':
                file_name = x
        # 如果没有.exp文件，抛出异常
        if file_name == '':
            error_message = 'Missing *.EXP file'
            create_error_message(self._output_excel_path, error_message)
            raise FileNotFoundError

        Mfg_idx = None
        Mfg_part_number_idx = None
        f7_idx = None
        no_mfg_component_list = []
        component_f7_dict = {}
        component_bom_dict = {}

        with open(os.path.join(dsn_path, file_name), 'r', encoding='gb18030') as file:
            id_idx, subsystem_idx, SSID_idx, bom_idx, description_idx, part_type_idx = None, None, None, None, None, None
            file.readline()
            title_list = file.readline().split('\t')
            for title_idx in range(len(title_list)):
                title_item = title_list[title_idx][1:-1]
                if title_item.upper() == 'PART REFERENCE':
                    part_reference_idx = title_idx
                elif title_item.upper() == 'ID':
                    id_idx = title_idx
                elif title_item.upper() == 'MFG':
                    Mfg_idx = title_idx
                elif title_item.upper() == 'MFG PART NUMBER':
                    Mfg_part_number_idx = title_idx
                elif title_item.upper() == 'F7':
                    f7_idx = title_idx
                # elif (title_item.upper().find('SUBSYSTEM') > -1 or
                #         title_list[title_idx].upper().find('SUB SYSTEM') > -1) and subsystem_idx is None:
                #     subsystem_idx = title_idx
                # elif title_item.upper() == 'SSID':
                #     subsystem_idx = title_idx
                elif title_item.upper() == 'BOM':
                    bom_idx = title_idx
                # elif title_item.upper() == 'DESCRIPTION':
                #     description_idx = title_idx
                # elif title_item.upper().find('PART TYPE') > -1:
                #     part_type_idx = title_idx
            # print(part_type_idx)
            if Mfg_idx is None or Mfg_part_number_idx is None:
                no_mfg_component_list.append(part_reference_idx)
            # 如果没有找到ID或Sub system则报错
            # if id_idx is None:
            #     error_message = 'There are no "Part Reference" in the properties'
            #     create_error_message(self._output_excel_path, error_message)
            #     raise FileNotFoundError
            #
            # if subsystem_idx is None:
            #     error_message = 'There is no "Sub system" or "SSID" in the properties'
            #     create_error_message(self._output_excel_path, error_message)
            #     raise FileNotFoundError

            # if bom_idx is None:

            #     error_message = 'There is no "BOM" in the properties'
            #     create_error_message(self._output_excel_path, error_message)
            #     raise FileNotFoundError

            # if description_idx is None and part_type_idx is None:
            #     error_message = 'There is no "Description" in the properties'
            #     create_error_message(self._output_excel_path, error_message)
            #     raise FileNotFoundError

            # if part_type_idx is None:
            #     error_message = 'There is no "Part Type" in the properties'
            #     create_error_message(self._output_excel_path, error_message)
            #     raise FileNotFoundError

            # id_bom_dict = {}
            # ssid_id_dict = {}
            # component_description_dict = {}
            # id_type_dict = {}

            for line in file.readlines():
                line_list = line.split('\t')
                line_id = line_list[id_idx][1:-1]
                line_part_reference = line_list[part_reference_idx][1:-1]

                if Mfg_idx and Mfg_part_number_idx:
                    line_mfg = line_list[Mfg_idx][1:-1]
                    line_mfg_part_number = line_list[Mfg_part_number_idx][1:-1]
                    component_mfg_dict[line_id] = line_mfg + ' : ' + line_mfg_part_number
                if f7_idx:
                    line_f7 = line_list[f7_idx][1:-1]
                    component_f7_dict[line_id] = line_f7
                if bom_idx:
                    line_bom = line_list[bom_idx][1:-1]
                    component_bom_dict[line_id] = line_bom
                # line_sub = line_list[subsystem_idx][1:-1]
                # if SSID_idx:
                #     line_ssid = line_list[SSID_idx][1:-1]
                # line_bom = line_list[bom_idx][1:-1]
                # line_description = line_list[description_idx][1:-1]
                # line_type = line_list[part_type_idx][1:-1]
                # if SSID_idx:
                #     ssid_id_dict[line_ssid] = ssid_id_dict.get(line_ssid, []) + [line_part_reference]
                # self.sub_component_dict_[line_sub] = self.sub_component_dict_.get(line_sub, []) + [line_part_reference]
                # id_bom_dict[line_part_reference] = id_bom_dict.get(line_part_reference, '') + line_bom
                # component_description_dict[line_part_reference] = component_description_dict.get(line_part_reference, '') + line_description
                # id_type_dict[line_part_reference] = id_type_dict.get(line_part_reference, '') + line_type

            component_NI_list = [key for key, value in component_bom_dict.items() if value == 'NI'] \
                if foxc_flag else [key for key, value in component_f7_dict.items() if value == '(R_)']
            return component_mfg_dict, component_NI_list

    def extract_bom_list(self, dsn_path):
        xlsx_name_list = [os.path.join(dsn_path, file) for file in os.listdir(dsn_path)
                          if file.find('.xlsx') > -1 and file.find('$') == -1]
        df = pd.read_excel(xlsx_name_list[0])
        mfg_list = [i.split(';') for i in list(df['Manufacturer'])]
        mfg_part_number_list = [i.split(';') for i in list(df['Manufacturer Part Number'])]

        mfg_mix_list = [mfg_item + ' : ' + mfg_num_item for Mfg_list, Mfg_num_list in
                        zip(mfg_list, mfg_part_number_list) for mfg_item, mfg_num_item in zip(Mfg_list, Mfg_num_list)]

        location_list = [i.split(',') for i in list(df['Location'])]
        len_list = [len(i) for i in mfg_list]

        mfg_mix_final_list = []
        len_idx = 0
        for i in len_list:
            next_idx = len_idx + i
            mfg_mix_final_list.append(' , '.join(mfg_mix_list[len_idx:next_idx]))
            len_idx = next_idx

        location_mfg_mix_dict = {location_item: mfg_mix_item for location_item_list, mfg_mix_item in
                                 zip(location_list, mfg_mix_final_list) for location_item in location_item_list}

        return location_mfg_mix_dict


class TraceDataProcessing:

    def __init__(self):
        self.extractIOData = ExtractIOData()
        self._input_path = self.extractIOData._root_path
        self._output_path = self.extractIOData._output_path
        self._output_excel_path = self.extractIOData._output_excel_path
        # pstxnet.dat
        self.all_net_list_ = []
        self.net_component_dict_ = {}
        self.net_component_list_ = []

        # pstxprt.dat
        self.all_component_list_ = []
        self.component_page_dict_ = {}
        self.primitive_name_component_name_dict_ = {}
        self.component_name_primitive_name_dict_ = {}

        # pstchip.dat
        self.primitive_name_pin_num_dict_ = {}
        self.primitive_name_pin_list_dict = {}

        # exp
        self.component_mfg_dict_ = {}
        self.extract_bom_dict_ = []
        self.component_NI_list_ = {}

        # fitt_net_connection_info
        self._pin_detail_info_dict = {}
        self._IC_list = {}

        # fit_connection_info_by_page
        self._IC_by_page_list = []
        self._connection_info_by_page_dict = {}

    # 输入文件夹名称，输出绝对路径
    def get_dsn_path(self, dir_name_list):

        dsn_path_list = [os.path.join(self._input_path, dir_name) for dir_name in dir_name_list]

        return dsn_path_list

    # 对所有文档资料进行数据处理
    def fit_all_dat_data(self, path):
        self.all_net_list_, self.net_component_list_, self.net_component_dict_ = \
            self.extractIOData.extract_pstxnet(path)
        self.all_component_list_, self.component_page_dict_, self.primitive_name_component_name_dict_, \
        self.component_name_primitive_name_dict_ = self.extractIOData.extract_pstxprt(path)
        self.primitive_name_pin_num_dict_, self.primitive_name_pin_list_dict = self.extractIOData.extract_pstchip(path)
        if path.find('foxconn') > -1:
            self.component_mfg_dict_, self.component_NI_list_ = self.extractIOData.extract_exp(path)

            return self.component_mfg_dict_
        # print(self.component_mfg_dict_)
        else:
            _, self.component_NI_list_ = self.extractIOData.extract_exp(path, foxc_flag=False)
            self.extract_bom_dict_ = self.extractIOData.extract_bom_list(path)
                    
            return self.extract_bom_dict_

    # 对每一个IC（pin num >= 3）跑出详细出pin走线信息
    # 形式（{ic: {pin: [[net, component, IC], [net, component, IC]]}）
    def fit_net_connection_info(self):
        Exclude_Net_List, PWR_Net_List, GND_Net_List = get_exclude_netlist(self.all_net_list_)
        IC_pin_num_dict = {}
        ic_pin_list_dict = {}
        for component_name in self.component_name_primitive_name_dict_.keys():
            primitive_name = self.component_name_primitive_name_dict_.get(component_name)
            IC_pin_num_dict[component_name] = self.primitive_name_pin_num_dict_[primitive_name]
            ic_pin_list_dict[component_name] = self.primitive_name_pin_list_dict[primitive_name]

        self._IC_list = [key for key, value in self.component_name_primitive_name_dict_.items()
                         if IC_pin_num_dict.get(key) > 2 and key not in self.component_NI_list_]
        # print('ic_pin_list_dict', ic_pin_list_dict)
        # print(IC_pin_num_dict)
        self.pin_detail_info_dict = get_detail_layout_info(self.net_component_list_, self._IC_list, ic_pin_list_dict,
                                                           IC_pin_num_dict, PWR_Net_List, GND_Net_List, self.component_NI_list_)

    def fit_connection_info_by_page(self, page):
        # print(self.component_page_dict_)
        components_by_page_list = [key for key, value in self.component_page_dict_.items() if value == str(page)]
        # print(components_by_page_list)
        self._IC_by_page_list = list(set(self._IC_list) & set(components_by_page_list))
        self._connection_info_by_page_dict = {ic: self.pin_detail_info_dict.get(ic) for ic in self._IC_by_page_list}

        return self._IC_by_page_list, self._connection_info_by_page_dict

# Function part


# 定义错误输出
def create_error_message(excel_path, error_message):
    # 創建excel
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet('error_message')

    error_format = workbook.add_format({'font_size': 22})

    worksheet.write('A1', 'Program running error:', error_format)
    worksheet.write('B2', error_message + ', please check and try again!', error_format)

    workbook.close()


# 将多维list展开成一维
def flatten(a):
    if not isinstance(a, (list,)) and not isinstance(a, (tuple,)):
        return [a]
    else:
        b = []
        for item in a:
            b += flatten(item)
    return b


# 自适应功能
def set_column_width(columns, worksheet):
    length_list = [ceil(max([len(str(y)) for y in x]) * 1.5) for x in columns]
    for i, width in enumerate(length_list):
        # print(i, i, width + 5)
        worksheet.set_column(i, i, width)


# 将信号线分为电源线与地线（电源线和地线的划分有待商榷，need discuss）

# '[0-9]V.*?_S[0-5]' 是电源线，U7202：OUT：5V_VCCPD_VBUS_F 是电源线，U7201：VCCD：3D3V_PD_VCCD 不是电源线
# Q7201：6：5V_USB_TYPEC_DIS_F 不是电源线，U7201：VDDIO：3D3V_VDDD 是电源线
def get_exclude_netlist(netlist):  # netlist = All_Net_List
    # Get pwr and gnd net list
    PWR_Net_KeyWord_List = ['^\+.*', '^-.*',
                            'VREF|PWR|VPP|VSS|VREG|VCORE|VCC|VT|VDD|VLED|PWM|VDIMM|VGT|VIN|[^S](VID)|VR',
                            'VOUT|VGG|VGPS|VNN|VOL|VSD|VSYS|VCM|VSA',
                            '.*\+[0-9]V.*', '.*\+[0-9][0-9]V.*', '\dV', '[0-9]V.*?_S[0-5]']
    PWR_Net_List = [net for net in netlist for keyword in PWR_Net_KeyWord_List if re.findall(keyword, net) != []]
    PWR_Net_List = sorted(list(set(PWR_Net_List)))

    GND_Net_List = [net for net in netlist if net.find('GND') > -1]
    GND_Net_List = sorted(list(set(GND_Net_List)))

    # 被排除的线：地线和电源线
    Exclude_Net_List = sorted(list(set(PWR_Net_List + GND_Net_List)))

    return Exclude_Net_List, PWR_Net_List, GND_Net_List


def create_pin_mapping_excel(output_path, IC_by_page_list_fox, IC_by_page_list_com, component_mfg_dict,
                             extract_bom_list):
    mfg_by_page_list = [component_mfg_dict.get(i) for i in IC_by_page_list_fox]
    bom_by_page_list = [extract_bom_list.get(i) for i in IC_by_page_list_com]

    IC_mfg_by_page_dict = {x: y for x, y in zip(IC_by_page_list_fox, mfg_by_page_list)}
    IC_bom_by_page_dict = {x: y for x, y in zip(IC_by_page_list_com, bom_by_page_list)}

    same_ic_com_fox_list = [[key_com, key_fox, '', ''] for key_fox, value_fox in IC_mfg_by_page_dict.items()
                            for key_com, value_com in IC_bom_by_page_dict.items() if value_fox in value_com]

    # 将数据填入表格
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1, 'border': True})
    border = workbook.add_format({'border': True})
    instruction_format = workbook.add_format({'color': 'red', 'font_size': 16})
    worksheet.write_row('A1', ['COMMON_IC', 'FOXCOON_IC', 'COMMON_PIN_LIST', 'FOXCOON_PIN_LIST'], bold)

    num = 0
    for i in same_ic_com_fox_list:
        num += 1
        worksheet.write_row(num, 0, i, border)

    worksheet.write(num + 2, 0, 'Instructions for use:', instruction_format)
    worksheet.write(num + 3, 0, '1. The first two columns are automatically genrated by the program. ',
                    instruction_format)
    worksheet.write(num + 4, 0, '2. The order of the pin lists filled in the last two columns needs to be consistent.',
                    instruction_format)

    set_column_width([['COMMON_IC'], ['FOXCOON_IC'], ['COMMON_PIN_LIST'], ['FOXCOON_PIN_LIST']], worksheet)
    workbook.close()


# Main part

# 对分类出的IC跑出详细走线信息（直流电源和地之间的滤波电容无法读出，如：U9302, need discuss）
def get_detail_layout_info(net_component_list, all_IC_list, ic_pin_list_dict, IC_pin_num_dict, Power_Net_List, GND_Net_List,
                           component_NI_list):
    """获取详细走线信息"""
    # print(net_component_list)
    ic_pin_component_dict = {}
    # 创建二维dict

    def addtwodimdict(thedict, key_a, key_b, val):
        if key_a in thedict.keys():
            thedict[key_a].update({key_b: val})
        else:
            thedict.update({key_a: {key_b: val}})
    # print('net_component_list', net_component_list)
    net_component_copy_list = copy.deepcopy(net_component_list)
    # 对每一个多pin芯片进行check
    for ic_name in all_IC_list:
        # print('ic_name', ic_name)
        # 找出每一个IC的所有pin脚
        try:
            ic_pin_list = ic_pin_list_dict[ic_name]
        except KeyError:
            ic_pin_list = ic_pin_list_dict[ic_name[:-1]]

        pin_net_component_list = []
        pin_net_component_dict = {}
        # print('ic_pin_list', ic_pin_list)
        # 对每一个pin name进行数据处理，找出与之对应的走线规则
        for pin_idx in range(len(ic_pin_list)):
            # 遍历pin name
            # 遍历没有拼写错误的所有pin脚
            if ic_pin_list[pin_idx]:
                # print(1, ic_pin_list[pin_idx])
                component_item_flag = False
                # final_flag = False
                no_nc_flag = False
                for component_item in net_component_list:
                    # print(net_component_list)
                    net_item = component_item[0]
                    # 找到pin脚连接的信号线信息
                    pin_find_idx_list = [i for i, v in enumerate(component_item[1:]) if v == ic_pin_list[pin_idx]]
                    ic_flag = False
                    if pin_find_idx_list:
                        for x in pin_find_idx_list:
                            if ic_name == component_item[1:][x - 1]:
                                ic_flag = True
                    if ic_flag:
                        pin_net_component_list1 = []
                        pin_net_component_list2 = []

                        component_item1 = copy.deepcopy(component_item)
                        component_item1.pop(component_item1.index(ic_pin_list[pin_idx], 1) - 1)
                        component_item1.pop(component_item1.index(ic_pin_list[pin_idx], 1))
                        component_item1 = component_item1[1::2]
                        flagfour = True
                        split_flag = True
                        split_component_list = []
                        layer_num = 0
                        layer_add_num_dict = {}

                        # 如果匹配到NC，因为NC在最后，所以匹配到NC说明前面都不匹配
                        if net_item == 'NC':
                            # 如果只匹配到NC
                            if no_nc_flag is False:
                                # if str(GPIO_pupd_list_org[pin_idx]) != 'None':
                                pin_net_component_list1 = []
                                pin_net_component_list2 = []
                                # print('NC', pin_net_component_list1)
                                pin_net_component_dict[ic_pin_list[pin_idx]] = [pin_net_component_list1]
                                component_item_flag = True
                        # 如果pin脚为 GND 或 VCC， 则不再往下走
                        elif net_item in Power_Net_List:
                            pin_net_component_dict[ic_pin_list[pin_idx]] = ['VCC']
                        elif net_item in GND_Net_List:
                            pin_net_component_dict[ic_pin_list[pin_idx]] = ['GND']
                        else:
                            no_nc_flag = True
                            # final_flag = False
                            while flagfour:
                                if pin_net_component_list1 and pin_net_component_list1[-1] == 'NC':
                                    pin_net_component_dict[ic_pin_list[pin_idx]].append(pin_net_component_list1)
                                    break
                                component_item3 = []

                                break_flag = False
                                split_out_flag = False
                                # if pin_name_list[pin_idx] == 'GPIO1':
                                #     print(1, component_item1)
                                if split_flag:
                                    split_component_list.append(copy.deepcopy(component_item1))
                                # 对未连接完成的线进行排除，未连接完成的线会无限向split_component_list中append([])
                                if len(split_component_list) > 50:
                                    break
                                # print('component_item1', component_item1)
                                for x_idx in range(len(component_item1)):
                                    # IC_flag = False
                                    next_flag = False
                                    all_break = False
                                    add_num = 0
                                    layer_num += 1

                                    item1 = component_item1[x_idx]
                                    # print('item1', item1)
                                    split_component_list[-1].pop(split_component_list[-1].index(component_item1[x_idx]))
                                    if 1 > 4:
                                        pass
                                    # 上件不上件的判断条件有待商榷，need discuss
                                    # # 对下一个经过的元器件是否是NI进行判断
                                    if item1 in component_NI_list:
                                        pin_net_component_list1.append('NI')
                                        pin_net_component_list2.append('NI')
                                        split_component_list.append([])
                                        split_flag = False
                                        add_num += 1
                                    # 如果上电
                                    else:
                                        # 大于4并且不是排阻则说明到另外一个芯片了，停止 item1 not in all_res_list and
                                        if IC_pin_num_dict[item1] >= 3:
                                            split_flag = False
                                            pin_net_component_list1.append(item1)
                                            pin_net_component_list2.append(item1)
                                            add_num += 1
                                            split_component_list.append([])
                                            # print(1, pin_net_component_list1)
                                        else:
                                            # print(pin_net_component_list1)
                                            # 判断是否为终止端元器件（中间的元器件会出现两次）
                                            if component_item1.count(item1) >= 1:
                                                # count = 0
                                                for component_item2 in net_component_copy_list:
                                                    # count += 1
                                                    add_sch_flag = False
                                                    # 找到元器件所连接的另一根线
                                                    # 如果这次经过的线与上次或第一次相同，则退出
                                                    if item1 in component_item2 and component_item2 != component_item \
                                                            and component_item2 != component_item3:

                                                        # 如果中间没有经过过这个元器件则进入循环
                                                        if component_item2[0] not in pin_net_component_list2:
                                                            # print('in')
                                                            add_sch_flag = True
                                                            pin_net_component_list1.append(item1)
                                                            pin_net_component_list2.append(item1)
                                                            # pin_net_component_list1.append(component_item2[0])
                                                            pin_net_component_list2.append(component_item2[0])
                                                            add_num += 2
                                                            if component_item2[0] != component_item[0] and component_item2 \
                                                                    != component_item3:
                                                                # print('3')
                                                                if component_item2[0] in Power_Net_List + GND_Net_List:
                                                                    split_component_list.append([])
                                                                    split_flag = False
                                                                    break

                                                                component_item1 = copy.deepcopy(component_item2)
                                                                component_item3 = copy.deepcopy(component_item2)

                                                                component_item1.pop(component_item1.index(item1) - 1)
                                                                component_item1.pop(component_item1.index(item1))

                                                                component_item1 = component_item1[1::2]

                                                                next_flag = True
                                                                split_flag = True
                                                                break_flag = True
                                                                break
                                                                # break 不要break，是因为可能元器件有超过两个pin，
                                                                # 要所有都遍历到, 虽然速度会变慢
                                                            else:
                                                                all_break = True

                                                    if net_component_copy_list[
                                                        -1] == component_item2 and add_sch_flag is \
                                                            False:
                                                        # print(4)
                                                        split_flag = False
                                                        add_num += 1
                                                        split_component_list.append([])
                                                        pin_net_component_list1.append(item1)
                                                        pin_net_component_list2.append(item1)

                                                        if pin_net_component_list1 not in pin_net_component_list:
                                                            pin_net_component_list.append(pin_net_component_list1)

                                                        if pin_net_component_dict.get(
                                                                ic_pin_list[pin_idx]):
                                                            pin_net_dict_list = pin_net_component_dict[
                                                                ic_pin_list[pin_idx]]
                                                            pin_net_dict_list.append(pin_net_component_list1)
                                                            pin_net_component_dict[ic_pin_list[
                                                                pin_idx]] = pin_net_dict_list
                                                        else:
                                                            pin_net_component_dict[ic_pin_list[
                                                                pin_idx]] = [pin_net_component_list1]
                                            # print(2, pin_net_component_list1)
                                    if all_break:
                                        pass
                                    else:
                                        # print(5, next_flag)
                                        layer_add_num_dict[layer_num] = add_num
                                        split_component_flag = True
                                        before_layer_num = 0
                                        if next_flag is False:
                                            if split_flag is False or component_item1[-1] == item1:
                                                split_flag = True
                                                split_component_list.pop(-1)
                                                before_layer_num = copy.deepcopy(layer_num)
                                                layer_num = len(split_component_list)
                                                if split_component_list:
                                                    try:
                                                        while not split_component_list[-1]:
                                                            split_component_list.pop(-1)
                                                            layer_num -= 1
                                                            # print('layer_num', layer_num)
                                                    except IndexError:
                                                        pass

                                                if split_component_list:
                                                    component_item1 = split_component_list[-1]
                                                    # if len(component_item1) == 1:
                                                    split_component_list.pop(-1)
                                                    # if component_item1[-1] == item1:
                                                    break_flag = True
                                                    split_component_flag = True

                                                    # print('split_component_flag', split_component_flag)
                                                else:
                                                    split_component_flag = False
                                                    flagfour = False
                                                    if component_item1[-1] == item1:
                                                        split_out_flag = True

                                                if pin_net_component_list1 not in pin_net_component_list:
                                                    pin_net_component_list.append(pin_net_component_list1)

                                                if 'NI' in pin_net_component_list1:
                                                    if pin_net_component_dict.get(ic_pin_list[pin_idx]):
                                                        pass
                                                    else:
                                                        pin_net_component_dict[
                                                            ic_pin_list[pin_idx]] = []
                                                else:
                                                    if pin_net_component_dict.get(ic_pin_list[pin_idx]):
                                                        pin_net_dict_list = pin_net_component_dict[
                                                            ic_pin_list[pin_idx]]
                                                        pin_net_dict_list.append(pin_net_component_list1)
                                                        pin_net_component_dict[ic_pin_list[pin_idx]] = \
                                                            pin_net_dict_list
                                                    else:
                                                        pin_net_component_dict[ic_pin_list[pin_idx]] = \
                                                            [pin_net_component_list1]

                                        if split_component_flag:
                                            if before_layer_num:
                                                for layer_idx in range(layer_num, before_layer_num + 1):
                                                    # print(layer_idx)
                                                    if layer_add_num_dict[layer_idx] != 0:
                                                        pin_net_component_list1 = pin_net_component_list1[:
                                                                                                -
                                                                                                layer_add_num_dict[
                                                                                                    layer_idx]]
                                                        pin_net_component_list2 = pin_net_component_list2[:
                                                                                                    -
                                                                                                    layer_add_num_dict[
                                                                                                        layer_idx]]

                                                    layer_add_num_dict.pop(layer_idx)
                                                layer_num -= 1
                                        # print(6, pin_net_component_list1)
                                        if break_flag:
                                            break

                                        if split_out_flag:
                                            break

                        # if component_item == net_component_list[-1] and final_flag:
                        #     component_item_flag = True
                        if component_item_flag:
                            break

        # print('pin_net_component_dict', pin_net_component_dict)
        # print('\n')
        for x, y in pin_net_component_dict.items():
            addtwodimdict(ic_pin_component_dict, ic_name, x, y)
    # print(ic_pin_component_dict)
    # 比较IC时，pin如何对应，同一个IC名称是否相同，need discuss
    return ic_pin_component_dict


# 对两个DSN档进行比较
def dsn_compare(component_mfg_dict, IC_by_page_list_fox, connection_info_by_page_dict_fox, extract_bom_list,
                IC_by_page_list_com, connection_info_by_page_dict_com, output_excel_path,input_pin_mapping_excel_path):
    mfg_by_page_list = [component_mfg_dict.get(i) for i in IC_by_page_list_fox]
    bom_by_page_list = [extract_bom_list.get(i) for i in IC_by_page_list_com]

    IC_mfg_by_page_dict = {x: y for x, y in zip(IC_by_page_list_fox, mfg_by_page_list)}
    IC_bom_by_page_dict = {x: y for x, y in zip(IC_by_page_list_com, bom_by_page_list)}

    same_ic_com_fox_dict = {key_com: key_fox for key_fox, value_fox in IC_mfg_by_page_dict.items()
                            for key_com, value_com in IC_bom_by_page_dict.items() if value_fox in value_com}

    com_miss_ic_list = list(set(IC_by_page_list_com) - set(same_ic_com_fox_dict.keys()))
    fox_over_ic_list = list(set(IC_by_page_list_fox) - set(same_ic_com_fox_dict.values()))

    # 从 Correspondence_between_pins.xlsx 中获取pin的对应关系
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(input_pin_mapping_excel_path)
    sht = wb.sheets[0]

    content = sht.range('A2').options(expand='table').value
    fox_ic_pin_dict = {}
    com_ic_pin_dict = {}
    num = 0
    for item_i in content:
        num += 1
        # 异常处理
        try:
            # 如果为空
            item_i[0] = item_i[0].strip()
            item_i[1] = item_i[1].strip()
            item_i[2] = item_i[2].strip()
            item_i[3] = item_i[3].strip()

            if '' in [item_i[0], item_i[1], item_i[2], item_i[3]]:
                wb.close()
                app.quit()
                print('Line %d in Correspondence_between_pins.xlsx has null value.Please check!' % num)
                os.system("pause")
                raise FileNotFoundError
        except:
            # 如果不能strip说明为None
            wb.close()
            app.quit()
            print('Line %d in Correspondence_between_pins.xlsx has null value.Please check!' % num)
            os.system("pause")
            raise FileNotFoundError

        com_ic_pin_dict[item_i[0]] = [i.strip() for i in item_i[2].split(',')]
        fox_ic_pin_dict[item_i[1]] = [i.strip() for i in item_i[3].split(',')]

    wb.close()
    app.quit()

    # print(same_ic_com_fox_dict)
    # print('')
    # print(IC_mfg_by_page_dict)
    # print('')
    # print(IC_bom_by_page_dict)
    # print('')
    # print(com_miss_ic_list)
    # print('')
    # print(fox_over_ic_list)
    # print('')

    # 创建输出excel
    workbook = xlsxwriter.Workbook(output_excel_path)
    worksheet = workbook.add_worksheet()
    # 添加用于突出显示单元格的粗体格式。
    # merge_format = workbook.add_format({'bold': True, 'font_size': 12})
    bold = workbook.add_format({'bold': True})
    border = workbook.add_format({'border': 1})

    worksheet.write('B1', 'Unmatchable pins', bold)
    worksheet.write('C1', 'Unmatched components', bold)
    row, col = 1, 0
    all_com_diff_pin_list = []
    # 相同IC两两比较
    for com_ic, fox_ic in same_ic_com_fox_dict.items():
        # 保存支路不同的pin
        com_diff_pin_list = []

        fox_pin_info_dict = connection_info_by_page_dict_fox.get(fox_ic)
        com_pin_info_dict = connection_info_by_page_dict_com.get(com_ic)

        fox_pin_net_num_dict = {key: len(value) for key, value in fox_pin_info_dict.items()}
        com_pin_net_num_dict = {key: len(value) for key, value in com_pin_info_dict.items()}

        # print('in1', ','.join(list(fox_pin_net_num_dict.keys())))
        # print('in2', ','.join(list(com_pin_net_num_dict.keys())))
        # print('')

        # 相同PIN两两比较
        for fox_pin, com_pin in zip(fox_ic_pin_dict[fox_ic], com_ic_pin_dict[com_ic]):
            # 如果pin的支路个数不同
            if fox_pin_net_num_dict[fox_pin] != com_pin_net_num_dict[com_pin]:
                com_diff_pin_list.append(com_pin)
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'})

        worksheet.merge_range(row, col, row + len(com_diff_pin_list) - 1, col, com_ic, merge_format)
        for item in com_diff_pin_list:
            worksheet.write(row, col + 1, item, border)
            row += 1
        all_com_diff_pin_list += com_diff_pin_list

        # print(com_ic)
        # print(com_pin_net_num_dict)
        # print('')
        # print(fox_ic)
        # print(fox_pin_net_num_dict)
        # print('')
        # print(com_pin_info_dict)
        # print('')
        # print(fox_pin_info_dict)
        # print('')
        # print(com_diff_pin_list)

    row, col = 1, 2
    for i in fox_over_ic_list:
        worksheet.write(row, col, i, border)
        row += 1

    # 设置自适应
    set_column_width([list(com_ic_pin_dict.keys()) + list(fox_ic_pin_dict.keys()), ['Unmatchable pins'] +
                      all_com_diff_pin_list, ['Unmatched components'] + fox_over_ic_list], worksheet)

    workbook.close()


def main():
    # 读取数据
    # foxconn 29
    foxconn_page = eval(input('Foxconn Page: '))
    # 24
    common_page = eval(input('Common Design Page: '))

    trace_data_processing_fox = TraceDataProcessing()
    input_pin_mapping_excel_path = os.path.join(trace_data_processing_fox._input_path,
                                                'Correspondence_between_pins.xlsx')
    output_pin_mapping_excel_path = os.path.join(trace_data_processing_fox._output_path,
                                                'Correspondence_between_pins.xlsx')
    my_file = Path(input_pin_mapping_excel_path)
    output_excel_path = os.path.join(trace_data_processing_fox._output_path, 'result.xlsx')
    dsn_path_list = trace_data_processing_fox.get_dsn_path(['foxconn', 'common'])
    component_mfg_dict = trace_data_processing_fox.fit_all_dat_data(dsn_path_list[0])
    trace_data_processing_fox.fit_net_connection_info()
    IC_by_page_list_fox, connection_info_by_page_dict_fox = \
        trace_data_processing_fox.fit_connection_info_by_page(foxconn_page)
    # print(IC_by_page_list_fox)
    # print('')
    # print(connection_info_by_page_dict_fox)
    # common
    trace_data_processing_com = TraceDataProcessing()
    extract_bom_list = trace_data_processing_com.fit_all_dat_data(dsn_path_list[1])
    trace_data_processing_com.fit_net_connection_info()
    IC_by_page_list_com, connection_info_by_page_dict_com = \
        trace_data_processing_com.fit_connection_info_by_page(common_page)

    # 如果'Correspondence_between_pins.xlsx'存在
    if my_file.exists():
        dsn_compare(component_mfg_dict, IC_by_page_list_fox, connection_info_by_page_dict_fox, extract_bom_list,
                    IC_by_page_list_com, connection_info_by_page_dict_com, output_excel_path,
                    input_pin_mapping_excel_path)
    # 如果不存在就生成
    else:
        create_pin_mapping_excel(output_pin_mapping_excel_path, IC_by_page_list_fox, IC_by_page_list_com,
                                 component_mfg_dict, extract_bom_list)
    # trace_data_processing.get_net_connection_info()
    # trace_data_processing.get_connection_info_by_page(29)
    # 客户的dsn文档
    # sub_component_dict1, pin_detail_info_dict1, category_components_dict1 = get_specific_ic_item(extractIOData, dsn1_path)
    # foxconn的dsn文档
    # sub_component_dict2, pin_detail_info_dict2, category_components_dict2 = get_specific_ic_item(extractIOData, dsn2_path)
    #
    # dsn_compare(sub_component_dict1, sub_component_dict3, pin_detail_info_dict1, pin_detail_info_dict3,
    #             category_components_dict1, category_components_dict3)


main()
