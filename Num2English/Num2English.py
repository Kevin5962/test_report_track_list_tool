'''
Description:
Utility for ENG convert the number part in asr to pure number
author: fengkai.xiao@cerence.com
date: 2022-03-29
'''


def Num2English_Normal(oldstr):
    dict_value ={'0': 'zero', '1': 'one', '2': 'two', '3': 'three','4': 'four','5': 'five', '6': 'six', '7': 'seven', '8': 'eight', '9': 'nine'}
    num_list =['0','1','2','3','4','5','6','7','8','9']
    exist_num =False
    for every_str in oldstr :
        if every_str in num_list:
            exist_num =True
            break
    if exist_num ==False:
        return oldstr
    oldstr_list =[]
    if exist_num == True:
        for every_str in oldstr:
            if every_str in num_list:
                oldstr_list.append(dict_value[every_str])
            else:
                oldstr_list.append(every_str)
        oldstr = ' '.join(oldstr_list)
        return oldstr


def get_number_part(oldstr):

    dict_value = {'0':'zero','1':'one','2':'two','3':'three','4':'four','5':'five','6':'six','7':'seven',\
    '8':'eight','9':'nine','10':'ten','11':'eleven','12':'twelve','13':'thirteen',\
    '14':'fourteen','15':'fifteen','16':'sixteen','17':'seventeen','18':'eighteen','19':'nineteen',\
    '20':'twenty','30':'thirty','40':'forty','50':'fifty','60':'sixty','70':'seventy','80':'eighty','90':'ninety',\
    '100':'hundred','200':'two hundred','300':'three hundred','400':'four hundred','500':'five hundred',\
    '600':'six hundred','700':'seven hundred','800':'eight hundred','900':'nine hundred','1000':'thousand','and':'and','.':'point','point':'.'}
    dict_value1 = {'0':'zero','1':'one','2':'two','3':'three','4':'four','5':'five','6':'six','7':'seven','8':'eight','9':'nine'}
    dict_value2 =  {'10':'ten','11':'eleven','12':'twelve','13':'thirteen','14':'fourteen','15':'fifteen',\
        '16':'sixteen','17':'seventeen','18':'eighteen','19':'nineteen'}
    dict_value3 = {'20':'twenty','30':'thirty','40':'forty','50':'fifty','60':'sixty','70':'seventy','80':'eighty','90':'ninety'}
    dict_value4 = {'100':'one hundred','200':'two hundred','300':'three hundred','400':'four hundred','500':'five hundred',\
    '600':'six hundred','700':'seven hundred','800':'eight hundred','900':'nine hundred','1000':'one thousand'}
    dict_symbol = {'+':'plus','-':'minus','x':'multiply','÷':'dividedby','=':'equals'}
    # 个位 可以连接 整十 百位 千位
    num_list1 = list(dict_value1.values())
    # 十位 只能独立计算
    num_list2 = list(dict_value2.values())
    # 十位 既可以独立计算 也可以连接个位一起计算
    num_list3 = list(dict_value3.values())
    # 可以进行一起计算的个位和十位list
    num_list4 = num_list1 + num_list3
    # 只占一个列表元素的个位和十位（就是排除百位和千位以后剩下的可能性）
    num_list5 = num_list1 + num_list2 + num_list3
    # 百位和千位
    num_list6 = list(dict_value4.keys())
    # 所有引入计算的字符串 包括数字和计算符号
    dict_value_total ={**dict_value,**dict_symbol}
    # 英文词汇 所有可供转换的阿拉伯数字 计算符号列表
    num_list_words = list(dict_value_total.values())
    # 阿拉伯数字 所有可供转换的阿拉伯数字 计算符号列表
    num_list_nums = list(dict_value_total.keys())
    # 英文的计算符号
    dict_symbol_list = list(dict_symbol.values())
    # 字典反转 key=英文 value=阿拉伯数字或计算符号
    inv_dict = {value:key for key,value in dict_value_total.items()}
    
    # 阿拉伯数字转化成 英文单词
    def num2words(oldstr):
        oldstr_list = oldstr.split(' ')
        # print('the origin oldstr_list is:')
        # print(oldstr_list)
        for i in range(len(oldstr_list)):
            i_str = oldstr_list[i]
            if i_str in num_list_nums:
                is_num = True
                oldstr_list[i] = dict_value_total[i_str]
        else:
            return ' '.join(oldstr_list)

    # 引入计算符号，加减乘除
    def cal_symbol(oldstr):
        newstr = oldstr
        newstr = newstr.lower().replace(' ','')
        # print('cal_symbol newstr is %s ' %newstr)
        is_symbol = False
        for i_symbol in dict_symbol_list:
            if i_symbol in newstr:
                # print(i_symbol)
                is_symbol = True
                break
        else:
            is_symbol = False

        if is_symbol:
            for i_symbol in dict_symbol_list:
                if i_symbol in newstr:
                    newstr = newstr.replace(i_symbol,inv_dict[i_symbol])
            # print(is_symbol,newstr)
            return newstr
        else:
            # print(is_symbol,oldstr)
            return oldstr

    def percent_nums(oldstr):
        # 百分数 %
        oldstr_list  = oldstr.split(' ')
        if 'percent' in oldstr_list:
            number_part = oldstr.replace('percent','%')
        else:
            number_part = oldstr
        return number_part

    def small_nums(raw_oldstr):
        # print('small number oldstr is %s ' %raw_oldstr)
        # 小数部分
        need_cal = False
        oldstr = ''
        point_part = ''
        oldstr_list = []
        int_oldstr_list = []
       
        oldstr = raw_oldstr.replace('point','.')
        # print('small number part oldstr is %s' %oldstr)
        if '.' in oldstr:
            need_cal = True
            # print('small number oldstr1 is %s ' %oldstr)
            # need_cal = True       
            oldstr_list = oldstr.split('.')
            # print('small part of oldstr_list is')
            # print(oldstr_list)
            int_oldstr_list = oldstr_list[0].split(' ')
            point_parts = oldstr_list[1].split(' ')
            # print('point_parts is')
            # print(point_parts)
            point_part = ''
            for i in point_parts:
                if i != '' and i in num_list1 or i.isdigit():
                    # print(i)
                    # print(inv_dict[str(i)])
                    if i.isdigit():
                        i_point_part = str(i)
                    else:
                        i_point_part = inv_dict[str(i)]
                    point_part += i_point_part
            point_part = '.' + point_part             
        else:
            need_cal = "double check"
            # print(need_cal)
            point_part = ''
            int_oldstr_list = oldstr.split(' ')

        return point_part,int_oldstr_list,need_cal

    def integer_parts(oldstr_list):
        # 整数部分
        four_parts = []
        Shi = Ge = 0
        if oldstr_list[-1] == '':
            del oldstr_list[-1]
        # print('int oldstr_list')
        len_oldstr_list = len(oldstr_list)
        if 'thousand' in oldstr_list:
            qian_index = oldstr_list.index('thousand') - 1
            Qian = inv_dict[oldstr_list[qian_index]]
        else:
            Qian = 0 
        four_parts.append(Qian)

        if 'hundred' in oldstr_list:
            bai_index = oldstr_list.index('hundred') - 1
            Bai = inv_dict[oldstr_list[bai_index]]
        else:
            Bai = 0   
        four_parts.append(Bai)

        if Qian !=0 or Bai != 0:
            if 'and' in oldstr_list:
                shi_index = oldstr_list.index('and') + 1
                if oldstr_list[shi_index].isdigit():
                    Shi = oldstr_list[shi_index]
                else:
                    Shi = inv_dict[oldstr_list[shi_index]] 
                if oldstr_list[-1] != oldstr_list[shi_index]:
                    if oldstr_list[-1].isdigit():
                        Ge = oldstr_list[-1]
                    else:
                        Ge = inv_dict[oldstr_list[-1]]
                else:
                    Ge = 0
            else:
                Shi = Ge = 0

        #存在千位和百位的时候，直接返回位数
            four_parts.append(Shi)
            four_parts.append(Ge)
            # print('four parts is')
            # print(four_parts)
            int_number_part = 1000*int(Qian) + 100*int(Bai) + int(Shi) + int(Ge)
            return str(int_number_part)
        else:
        # 只存在十位和个位的时候，直接返回对应的字符串
        # 1.对于只剩下个位和十位，对于出现 2.6/95 这种 无法一次性转化为单词的场景
        # 2.存在95这样的非 num_list_words 但是 阿拉伯数字
            oldstr = num2words(' '.join(oldstr_list))
            oldstr_list = oldstr.split()
            # print('the shi and ge oldstr_list:')
            # print(oldstr_list)
            len_oldstr_list = len(oldstr_list)
            # print('the num2words str: %s' %oldstr)
            if len_oldstr_list == 1:
                new_shi_str = ''
                if oldstr_list[0] in num_list_words:
                    new_shi_str = str(inv_dict[oldstr_list[0]])
                elif oldstr_list[0].isdigit(): 
                     new_shi_str = oldstr_list[0]
                return new_shi_str
            else:
                # 分段的数字 字符串
                cal_flag = ('0','1')
                new_flags_dict = {}
                new_flags2_dict = {}
                flag_values_list = []
                for i in range(len_oldstr_list):
                    i_str = oldstr_list[i]
                    if i_str == 'and':
                        new_flags_dict[i] = ['3',i_str]     # 中间的连接词
                    elif i_str in num_list2:
                        new_flags_dict[i] = ['2',i_str]     # 十位 只能独立计算
                    elif i_str in num_list1:
                        new_flags_dict[i] = ['1',i_str]     # 个位 可以合并计算
                    elif i_str in num_list3:
                        new_flags_dict[i] = ['0',i_str]     # 整十位 可以合并计算
                    else:
                        new_flags_dict[i] = ['4',i_str]     # 不包含 可计算的两位数字
                
                flag_values_list = list(new_flags_dict.values())

                for i in range(len_oldstr_list):
                    new_flags2_dict[i] = flag_values_list[i][0]

                # print('new_flags dict:')
                # print(new_flags_dict)
                # print('new_flags2_dict:')
                # print(new_flags2_dict)
                # print('flag_values_list:')
                # print(flag_values_list)

                new_cal_flags = []    
                for i in range(len(flag_values_list)-1):
                    i_str = new_flags2_dict[i]
                    j_str = new_flags2_dict[i+1]
                    new_cal_flags.append((i_str,j_str))

                flags_index = []
                for i in range(len(new_cal_flags)):
                    if new_cal_flags[i]  == cal_flag:
                        flags_index.append(i)

                many_num_str = ''
                flag_cal_index = []
                for i in range(len_oldstr_list):
                    for j in flags_index:
                        if i == j:                            
                            flag_cal_index.append(i)
                            flag_cal_index.append(i+1)

                # print('flag_cal_index:')
                # print(flag_cal_index)

                for i in range(len_oldstr_list):
                    if i in flag_cal_index and flag_cal_index.index(i)%2==0:
                        # print('the i_str :%s' %oldstr_list[i])
                        many_num_str += ' ' + str(int(inv_dict[oldstr_list[i]]) + int(inv_dict[oldstr_list[i+1]]))
                    elif i not in flag_cal_index:
                        if oldstr_list[i] in num_list_words:
                            many_num_str += ' ' + str(inv_dict[oldstr_list[i]])
                        else:
                            many_num_str += ' ' + oldstr_list[i]
                return many_num_str

    def single_nums(oldstr_list):
        # 全部是间隔数字
        # print('single')
        number_part_list =[]
        for every_str in oldstr_list:
            if every_str in num_list1:
                number_part_list.append(inv_dict[every_str])
            else:
                number_part_list.append(every_str)
        number_part = ' '.join(number_part_list)
        return number_part

    def is_num_part(oldstr):
        # 引入计算流程前 对字符串进行统一英文处理  
        oldstr = num2words(oldstr)
        # print('cal_origin:%s' %oldstr)
        # 判断数字是否需要计算,只有一段数字字符计算的全部流程
        number_part = ''
        point_part,oldstr_list,need_cal = small_nums(oldstr)
        # print('point_part:%s' %point_part)
        # print('small_nums return is')
        # print(small_nums(oldstr))
        if need_cal == True:
            # print('int oldstr_list')
            # print(oldstr_list)
            int_number_part = integer_parts(oldstr_list)
            number_part = int_number_part + point_part

            # print('int number is %s' %int_number_part)
            # print('small number is %s' %point_part)
            # print('number_part is %s' %number_part)

            return need_cal,number_part
        elif need_cal == "double check":
            # print(need_cal)
            # print(oldstr_list)
            for i in oldstr_list:
                if i in num_list_words and i not in num_list1:
                    need_cal = True
                    int_number_part = integer_parts(oldstr_list)
                    number_part = int_number_part
                    return need_cal,number_part
            else:
                need_cal = False
                number_part = single_nums(oldstr_list)
                return need_cal,number_part

    from itertools import groupby
    def series_list_split(xlist):
        fun = lambda x: x[1]-x[0]
        all_phases = []
        for k, g in groupby(enumerate(xlist), fun):
            scop = [j for i, j in g] 
            all_phases.append(scop) 
        # print("连续数字列表切分：{}".format(all_phases))
        return all_phases

    def many_number_parts(oldstr):
        oldstr_list = oldstr.split(' ')
        # print('oldstr_list is:')
        # print(oldstr_list)
        # 把整千和整百的阿拉伯数字换算成英文
        for i in range(len(oldstr_list)):
            i_str = oldstr_list[i]
            if i_str in list(dict_value4.keys()):
                oldstr_list[i] = dict_value4[i_str]

        oldstr_list = ' '.join(oldstr_list).split(' ')
        # print('the origin oldstr is:%s' %oldstr_list)

        common_str_index = []
        number_str_index = []
        newstr_list = []

        for i in range(len(oldstr_list)):
            i_str = oldstr_list[i]
            if i_str in num_list_words or i_str.replace('.','').isdigit() == True:
                number_str_index.append(i)
                newstr_list.append('')
            else: 
                common_str_index.append(i)
                newstr_list.append(i_str)

        # print(newstr_list)
        # print(common_str_index)
        # print(number_str_index)
        str_list = series_list_split(common_str_index)
        numstr_index_list = series_list_split(number_str_index)
        # print('str_list is:')
        # print(str_list)
        # print('numstr_index_list is:')
        # print(numstr_index_list)
        many_num_parts = dict()
        for numstr_i in numstr_index_list:
            # print('num_str_i is:')
            # print(numstr_i)
            if len(numstr_i) > 1:
                numstr_left =numstr_i[0]
                numstr_right = numstr_i[-1]+1  
                # print('give the many num_part_str is:')
                # print(' '.join(oldstr_list[numstr_left:numstr_right]))
                i_numstr = is_num_part(' '.join(oldstr_list[numstr_left:numstr_right]))[1]
            else:
                # print('give the single num_part_str is:')
                # print(oldstr_list[numstr_i[0]] +' ')
                i_numstr = is_num_part(oldstr_list[numstr_i[0]] +' ')[1]
            many_num_parts[numstr_i[0]] = i_numstr

        for xi in range(len(oldstr_list)):
            for left_index in many_num_parts.keys():
                if xi == left_index:
                    # print('xi = %s and the many_num_parts = %s' %(xi,many_num_parts[xi]))
                    newstr_list[xi] = many_num_parts[xi]
        # point_part = small_nums(oldstr)[0]
        # oldstr_list = oldstr_list + list(point_part)
        # print('newstr_list2 is')
        # print(newstr_list)
        newstr_list = [x for x in newstr_list if x!= '']
        # print('newstr_list3 is')
        # print(newstr_list)
        number_part1 = ' '.join(newstr_list)
        # print(number_part1)
        number_part2 = percent_nums(number_part1)
        # print(number_part1)
        number_part3 = cal_symbol(number_part2)
        # print(number_part3)
        return number_part3
    return  many_number_parts(oldstr)  
# 支持切换三位英文表达（包含个，十，百）转换到数字，及常规转换
def English2Num(oldstr):

    dict_value = {'0':'zero','1':'one','2':'two','3':'three','4':'four','5':'five','6':'six','7':'seven',\
    '8':'eight','9':'nine','10':'ten','11':'eleven','12':'twelve','13':'thirteen',\
    '14':'fourteen','15':'fifteen','16':'sixteen','17':'seventeen','18':'eighteen','19':'nineteen',\
    '20':'twenty','30':'thirty','40':'forty','50':'fifty','60':'sixty','70':'seventy','80':'eighty','90':'ninety',\
    '100':'hundred','200':'two hundred','300':'three hundred','400':'four hundred','500':'five hundred',\
    '600':'six hundred','700':'seven hundred','800':'eight hundred','900':'nine hundred','1000':'thousand','and':'and','.':'point','point':'.'}
    dict_symbol = {'+':'plus','-':'minus','x':'multiply','÷':'dividedby','=':'equals'}
    num_list = list(dict_value.values()) + list(dict_symbol.values()) + list(dict_symbol.keys())

    # print(num_list)

    def is_numbers(oldstr,num_list):
        # 判断是否存在数字
        exist_num = False
        oldstr_list = oldstr.split(' ')
        # print('the begin oldstr_list:')
        # print(oldstr_list)
        for every_str in oldstr_list :
            # 英文数字和计算符号
            if every_str in num_list:
                exist_num =True
                break
        else:
            # 阿拉伯数字
            for i_str in oldstr_list:
                if i_str.replace('.','').isdigit():
                    exist_num =True
                    break
            else:
                exist_num = False
        # print('exist_num %s' %str(exist_num))
        return exist_num

    # 符号替换
    def ignore_special_symbol(oldstr,exist_num):
        if exist_num == False and '-' in oldstr and oldstr[oldstr.index('-')+1].isdigit() == False:
            repalce_none_list = [',','-','?']
        else:
            repalce_none_list = [',','?']

        for i in repalce_none_list: 
            if i in oldstr:
                oldstr = oldstr.replace(i, ' ')
        return oldstr
    # 单位换算
    def unit_convert(oldstr):
        dict_units = {'degrees':'°','milesperhour':'mph','kilometersperhour':'km/h'}
        need_re_strs = list(dict_units.keys())
        unit_str = oldstr.replace(' ','')
        # print(unit_str)
        for i in need_re_strs:
            if i in unit_str:
                exist_unit = True
                break
        else:
            exist_unit = False
        if exist_unit:
            for i in need_re_strs:
                if i in unit_str: 
                    unit_str = unit_str.replace(i,dict_units[i])
            return unit_str
        else:
            return oldstr

    exist_num = is_numbers(oldstr,num_list)
    # print('exist_num:%s' %str(exist_num))
    oldstr = ignore_special_symbol(oldstr,exist_num)
    if exist_num:
        number_part_tmp = get_number_part(oldstr)
        # print('number_part_tmp is: %s' %number_part_tmp)
        number_part = unit_convert(number_part_tmp)
    else:
        number_part = oldstr
    return number_part

if __name__=='__main__':
    # 现在是 数字部分的str 在一个oldstr里面存在一个和多个怎样分类解决的问题
    # s1 = 'play FM one hundred . one two three'
    # print(English2Num(s1))
    # s2 = 'FM ninety five . five'
    # print(English2Num(s2))
    # s2 = 'FM 95.5'
    # print(English2Num(s2))
    # s2 = 'Play FM 100 and 2.6'
    # print(English2Num(s2))
    # s3 = 'Drive me to Philadelphia Spring Garden St thirty one'
    # print(English2Num(s3))   
    # s5 = 'page eight'
    # print(English2Num(s5))
    # s5 = 'channel five'
    # print(English2Num(s5))
    # s5 = 'channel one five'
    # print(English2Num(s5))
    # s6 = 'help me dail the phone number start with three nine five five eight eight'
    # print(English2Num(s6))
    # s7 = 'call hash one seven two star one five three'
    # print(English2Num(s7))
    # s8 = 'call hash one hundred and seventy two star one hundred and fifty three'
    # print(English2Num(s8))
    # s9 = 'call ash one hundred and seventy two Stunna one hundred and fifty three'
    # print(English2Num(s9))
    # s10 = '7 0 3 9 5 9 4 7'
    # print(English2Num(s10))
    # s10 = 'seven zero three nine five nine for seven'
    # print(English2Num(s10))
    # s10 = 'str1 str2 seven zero three str3 nine five nine four seven'
    # print(English2Num(s10))
    # s11 = 'str1 seven zero str2 str3 three nine five str4 nine four seven'
    # print(English2Num(s11))
    # s11 = 'str1 str2 seven zero str3 str4 three nine str5 str6 five nine seven'
    # print(English2Num(s11))
    # s12 = 'str1 str2 one hundred and seventy three str3 str4 three nine str5 five nine str6 twenty six'
    # print(English2Num(s12))
    # s13 = 'Broadcast AM 600 and three'
    # print(English2Num(s13))
    # s13 = 'Broadcast AM 600 and twenty three'
    # print(English2Num(s13))
    # s14 = 'Broadcast AM 1000 and twenty'
    # print(English2Num(s14))
    # s15 = 'The next one'
    # print(English2Num(s15))
    # s15 = 'activate Olson text and I do that strengths two'
    # print(English2Num(s15))
    # s15 = '''Show p.m. to don't five outside liquor'''
    # print(English2Num(s15))
    # s15 = 'Please turn up the volume for me by 20 percent'
    # print(English2Num(s15))
    # s15 = 'Please turn up the volume for me by twenty percent'
    # print(English2Num(s15))
    # s15 = 'play game of thrones season six for me'
    # print(English2Num(s15))
    # s15 = 'play game of thrones season six for me'
    # print(English2Num(s15))
    # s15 = 'Listen to The Last one from Still Night Still Light'
    # print(English2Num(s15))
    # s15 = ' Listen to the last one from still night still light'
    # print(English2Num(s15))
    # s15 = 'Open the window half way'
    # print(English2Num(s15))
    # s15 = 'Wi-Fi off'
    # print(English2Num(s15))
    # s15 = 'the speed reduced to one hundred kilometers per hour'
    # print(English2Num(s15))
    # s15 = '''give a call to Mary's phone number starting with one three eight'''
    # print(English2Num(s15))
    # s15 = 'play FM one hundred and two point six'
    # print(English2Num(s15))
    # s15 = 'How much is 345 minus 321'
    # print(English2Num(s15))
    # s15 = 'How much is 345-321'
    # print(English2Num(s15))
    # s15 = 'How much is 345x321-20÷5'
    # print(English2Num(s15))
    # s15 = 'How much is 345 multiply 321 minus 20 divided by 5'
    # print(English2Num(s15))
    # s15 = 'Switch on the rear-end collision warning system warning tone'
    # print(English2Num(s15))
    # s15 = 'xiaohongshu'
    # print(English2Num(s15))  
    # s4 = 'Drive me to Philadelphia Spring Garden St thirty one twelve thirty one twelve twelve thirty one'
    # print(English2Num(s4))
    # s4 = 'Drive me to Philadelphia Spring Garden St thirty one twelve 31 twelve twelve thirty one'
    # print(English2Num(s4))
    # s5 = 'Drive me to Philadelphia Spring Garden St thirty one twelve thirty and one twelve twelve twelve one'
    # print(English2Num(s5))
    # s5 = 'Turn on north and six on my strength one'
    # print(English2Num(s5))
    # s5 = 'and and six and six and six my one'
    # print(English2Num(s5))
    # s5 = 'the speed reduced to one hundred kilometers per hour'
    # print(English2Num(s5))
    # s= '20 miles per hour'
    # s= '20 kilometers per hour'
    # s= '20 mph'
    # s= '20 km/h'
    # s= '20'
    # s= '2.6'
    s= 'Correct 1250 to 1350'
    # s= 'Raise the temperature to 25 degrees'
    print(English2Num(s))