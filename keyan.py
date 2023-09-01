import xlrd
import xlwt
import random
import copy


class Gene():
    def __init__(self, maindata):
        self.gene1 = [0, 0, 0, 0, 0]   # 省物资0，省魔方1，缺心智2，金：彩3，彩：装备4
        self.gene2 = [0, 0, 0, 0, 0]   # 省物资5，省魔方6，缺心智7，金：彩8，彩：装备9
        self.gene3 = [0, 0, 0, 0, 0]   # 省物资10，省魔方11，缺心智12，金：彩13，彩：装备14
        self.changenum = [1, 2, 4, 8, 16, 32, -1, -2, -4, -8, -16, -32]
        self.max = [1024, 1024, 64, 256, 256]
        self.wuzi_work = 1
        self.mofang_work = 1
        self.xinzhi_work = 1
        self.canchange = []
        if maindata[31][1] == 0:
            self.wuzi_work = 0
        if maindata[32][1] == 0:
            self.mofang_work = 0
        if maindata[36][1] == 0:
            self.xinzhi_work = 0
        for i in range(15):
            if i == 0 or i == 5 or i == 10:
                if self.wuzi_work == 0:
                    continue
            if i == 1 or i == 6 or i == 11:
                if self.mofang_work == 0:
                    continue
            if i == 2 or i == 7 or i == 12:
                if self.xinzhi_work == 0:
                    continue
            if i == 8 or i == 13 or i == 14:
                continue
            self.canchange.append(copy.deepcopy(i))
        self.canchange_num = len(self.canchange) - 1

    def newrand(self):
        for i in range(5):
            if i == 0 and self.wuzi_work == 0:
                continue
            if i == 1 and self.mofang_work == 0:
                continue
            if i == 2 and self.xinzhi_work == 0:
                continue
            self.gene1[i] = random.randint(0, self.max[i])
            if i == 3:
                continue
            self.gene2[i] = random.randint(0, self.max[i])
            if i == 4:
                continue
            self.gene3[i] = random.randint(0, self.max[i])

    def change(self):
        p = random.randint(0, self.canchange_num)
        p = self.canchange[p]
        q = random.randint(0, 11)
        if p < 5:
            self.gene1[p] = self.changenum[q] + self.gene1[p]
            if self.gene1[p] < 0:
                self.gene1[p] = 0
        elif p < 10:
            p -= 5
            self.gene2[p] = self.changenum[q] + self.gene2[p]
            if self.gene2[p] < 0:
                self.gene2[p] = 0
        else:
            p -= 10
            self.gene3[p] = self.changenum[q] + self.gene3[p]
            if self.gene3[p] < 0:
                self.gene3[p] = 0


def excel(fname, sname):       # 读表函数
    wb = xlrd.open_workbook(fname)   # 打开Excel文件
    sheet = wb.sheet_by_name(sname)   # 通过excel表格名称(rank)获取工作表
    dat = []     # 创建空list
    for a in range(sheet.nrows):  # 循环读取表格内容（每次读取一行数据）
                cells = sheet.row_values(a)  # 每行数据赋值给cells
                data = cells      # 因为表内可能存在多列数据，0代表第一列数据，1代表第二列，以此类推
                dat.append(data)     # 把每次循环读取的数据插入到list
    return dat


def shouyi(values, maindata, jieduan):   # 给出当前倾向、项目价值下的最优参考投入产出
    valuelist = [[0, 0, 0, 0, 0] for i in range(29)]   # 消耗数值，产出数值，点击收益，刷出概率,项目消耗时间
    ranklist = [i for i in range(29)]
    value_wuzi = values[0] / 200
    value_mofang = values[1] * 10
    value_jintu = 1000
    value_caitu = value_jintu * values[3]
    value_caizhuang = value_caitu * values[4]
    value_time = 1500
    value_time2 = 0
    value_xinzhi = values[2] * 0.6
    guolv = 5
    xinjiabi = 0   # 性价比(未迭代)
    if jieduan == 1:
        value_jintu = 0
        j1 = maindata[10][2]
        maindata[10][2] = 0
        maindata[14][2] += j1
        j1 = maindata[11][2]
        maindata[11][2] = 0
        maindata[15][2] += j1
        j1 = maindata[12][2]
        maindata[12][2] = 0
        maindata[16][2] += j1
        j1 = maindata[13][2]
        maindata[13][2] = 0
        maindata[17][2] += j1

    if jieduan == 2:
        value_jintu = 0
        value_caitu = 1
        value_caizhuang = 100000
    p = 0
    while(1):
        for i in range(29):
            list_row = i + 2
            costs = maindata[list_row][3] * value_wuzi  # 物资消耗
            costs += maindata[list_row][4] * value_mofang  # 魔方消耗
            costs += maindata[list_row][1] * value_time + value_time2  # 时间消耗
            valuelist[i][0] = costs
            products = maindata[list_row][1] * maindata[list_row][5] * value_jintu  # 金图产出
            products += maindata[list_row][1] * maindata[list_row][6] * value_caitu  # 彩图产出
            products += maindata[list_row][1] * maindata[list_row][7] * value_caizhuang  # 彩装备产出
            products += maindata[list_row][1] * maindata[list_row][8] * value_xinzhi  # 心智产出
            valuelist[i][1] = products
            shuachu = maindata[list_row][2]  # 刷出概率
            valuelist[i][3] = shuachu
            valuelist[i][4] = maindata[list_row][1]
        for i in range(29):
            valuelist[i][2] = valuelist[i][1] - valuelist[i][0] * xinjiabi
        for i in range(29):  # 冒泡排序
            for j in range(28):
                rank1 = ranklist[j]
                rank2 = ranklist[j + 1]
                if valuelist[rank1][2] < valuelist[rank2][2]:
                    ranklist[j] = rank2
                    ranklist[j + 1] = rank1
        xinjiabi1 = shouyi_yuce(ranklist, valuelist, guolv, maindata)
        p += 1
        if p > 20:
            break
        if xinjiabi1[0] == xinjiabi:   # 迭代计算性价比期望，给出排序修正和参考投入产出
            break
        else:
            xinjiabi = xinjiabi1[0]
    while(1):
        guolv += 1
        xinjiabi2 = shouyi_yuce(ranklist, valuelist, guolv, maindata)
        if xinjiabi1[1] < xinjiabi2[1]:
            xinjiabi1 = copy.deepcopy(xinjiabi2)
        else:
            guolv -= 1
            break

    chanchu = xinjiabi1[2]
    chanchu.append(guolv)
    chanchu.append(ranklist)
    return chanchu


def shouyi_yuce(ranklist, valuelist, guolv, maindata):
    p_spare = 1     # 剩余刷出概率
    p_dproject = maindata[37][1]
    choose_spare = 1   # 剩余选取比例
    p_willchoose = 0
    cost_time = 0
    choose_p_list = []
    cost_all = 0
    value_all = 0
    for i in range(29):
        i1 = ranklist[i]  # 项目序号
        p0 = valuelist[i1][3]
        p1 = p0/p_spare
        p_spare -= p0
        if (i1 >= 8)*(i1 <= 15):
            p2 = p1 * p_dproject
        else:
            p2 = p1 / 4
        p3 = 1 - (1 - p1) * (1 - p1) * (1 - p1) * (1 - p2) * (1 - p2)
        p_choose = p3 * choose_spare
        choose_spare -= p_choose
        choose_p_list.append([p_choose,0])
        cost_time += p_choose * valuelist[i1][4]
        if i <= guolv:
            p_willchoose += p_choose
        cost_all += valuelist[i1][0]*p_choose
        value_all += valuelist[i1][1] * p_choose

    xinjiabi = [0, 0]
    xinjiabi[0] = value_all / cost_all
    cishu = 24/cost_time
    p_willnotchoose = 1-p_willchoose
    p_usefree = 1-pow(p_willchoose,cishu)
    p_false = (p_willnotchoose*cishu-p_usefree*p_willchoose)/cishu
    p_ture = 1 - p_false
    cost_all = 0
    value_all = 0
    daily_data = [0, 0, 0, 0, 0, 0, 0]  # 0物资，1魔方，2间隔，3金图，4彩图，5彩装，6心智
    cost_time = 0
    for i in range(29):  # 重新计算选取率
        i1 = ranklist[i]  # 项目序号
        if i <= guolv:
            p_choose = choose_p_list[i][0]
            p_choose_true = p_choose/p_willchoose
            p_choose_true *= p_ture
        else:
            p_choose = choose_p_list[i][0]
            p_choose_true = p_choose / p_willnotchoose
            p_choose_true *= p_false
        choose_p_list[i][1] = p_choose_true
        cost_all += valuelist[i1][0] * p_choose_true
        value_all += valuelist[i1][1] * p_choose_true
        cost_time += p_choose_true * valuelist[i1][4]
        list_row = i1 + 2
        daily_data[0] += maindata[list_row][3] * p_choose_true  # 每日物资
        daily_data[1] += maindata[list_row][4] * p_choose_true  # 每日魔方
        daily_data[2] += maindata[list_row][1] * p_choose_true  # 每日耗时
        daily_data[3] += maindata[list_row][1] * maindata[list_row][5] * p_choose_true  # 每日金图
        daily_data[4] += maindata[list_row][1] * maindata[list_row][6] * p_choose_true  # 每日彩图
        daily_data[5] += maindata[list_row][1] * maindata[list_row][7] * p_choose_true  # 每日彩装
        daily_data[6] += maindata[list_row][1] * maindata[list_row][8] * p_choose_true  # 每日心智
    cishu = 24 / cost_time
    for i in range(7):
        if i == 2:
            daily_data[i] *= 5
        else:
            daily_data[i] *= cishu
    xinjiabi[1] = value_all / cost_all
    xinjiabi.append(daily_data)
    return xinjiabi

def chanchulist(chan):
    global maindata
    need_g = 1029 - int(maindata[38][1])
    need_c = 1026 - int(maindata[39][1])
    day1 = need_g  / chan[0][3]  # 金船毕业时间
    cai = day1 * chan[0][4]  # 彩船在金船毕业时的数目
    day2 = 0  # 彩船毕业时间
    if cai < need_c:
        day2 = (need_c - cai) / chan[1][4]
    zhuangs = 0
    xinzhis = 0
    wuzi = 0
    mofang = 0
    if day1 <= 365:
        zhuangs += day1 * chan[0][5]
        xinzhis += day1 * chan[0][6]
        wuzi += day1 * chan[0][0]
        mofang += day1 * chan[0][1]
        if (day1 + day2) <= 365:
            zhuangs += day2 * chan[1][5]
            zhuangs += (365 - day1 - day2) * chan[2][5]
            xinzhis += day2 * chan[1][6]
            xinzhis += (365 - day1 - day2) * chan[2][6]
            wuzi += day2 * chan[1][0]
            wuzi += (365 - day1 - day2) * chan[2][0]
            mofang += day2 * chan[1][1]
            mofang += (365 - day1 - day2) * chan[2][1]

        else:
            zhuangs += (365 - day1) * chan[1][5]
            xinzhis += (365 - day1) * chan[1][6]
            wuzi += (365 - day1) * chan[1][0]
            mofang += (365 - day1) * chan[1][1]
    else:
        zhuangs += 365 * chan[0][5]
        xinzhis += 365 * chan[0][6]
        wuzi += 365 * chan[0][0]
        mofang += 365 * chan[0][1]
    wuzi_day = wuzi / 365
    mofang_day = mofang / 365
    day2 += day1
    return day1, day2, zhuangs, xinzhis, wuzi_day, mofang_day


def shiyingdu(maindata,  chan):       #当前倾向对产出的匹配适应度
    limit_wuzi = maindata[31][1]           # 物资限制方式
    limit_wuzi_num = maindata[31][2]       # 物资限制数目
    limit_mofang = maindata[32][1]         # 魔方限制方式
    limit_mofang_num = maindata[32][2]     # 物资限制数目
    limit_jinchuan = maindata[33][1]       # 金船限制方式
    limit_jinchuan_num = maindata[33][2]   # 金船限制数目
    limit_caichuan = maindata[34][1]       # 彩船限制方式
    limit_caichuan_num = maindata[34][2]   # 彩船限制数目
    limit_caizhuang = maindata[35][1]      # 彩装限制方式
    limit_caizhuang_num = maindata[35][2]  # 彩装限制数目
    limit_xinzhi = maindata[36][1]         # 心智限制方式
    limit_xinzhi_num = maindata[36][2]     # 心智限制数目
    limit_caizhuang_num *= 50
    fen = 100
    day1, day2, zhuangs, xinzhis, wuzi_day, mofang_day = chanchulist(chan)
    if limit_jinchuan == 2:    # 金船限制为尽量快
        fen += (365 - day1)
    if limit_caichuan == 2:
        fen += (365 - day2)
    if limit_caizhuang == 2:
        fen += zhuangs
    if limit_xinzhi == 2:
        fen += xinzhis / 180
    if limit_wuzi == 2:
        fen += (40000 - wuzi_day)/70
    if limit_mofang == 2:
        fen += (60 - mofang_day)*6
    if limit_wuzi == 1:
        if limit_wuzi_num < wuzi_day:
            fen /= (1 + (wuzi_day- limit_wuzi_num) * 0.001)            # 不满足物资限制扣分
    if limit_mofang == 1:
        if limit_mofang_num < mofang_day:
            fen /= (1 + (mofang_day - limit_mofang_num) * 1)      # 不满足魔方限制扣分
    if limit_jinchuan == 1:
        if limit_jinchuan_num < day1:
            fen /= (1 + (day1 - limit_jinchuan_num) * 20)      # 不满足金船毕业时间限制扣分
    if limit_caichuan == 1:
        if limit_caichuan_num < day2:
            fen /= (1 + (day2 - limit_caichuan_num) * 30)      # 不满足彩船毕业时间限制扣分
    if limit_caizhuang == 1:
        if limit_caizhuang_num > zhuangs:
            fen /= (1 + (limit_caizhuang_num - zhuangs) * 400)   # 不满装备毕业数目限制扣分
    if limit_xinzhi == 1:
        if limit_xinzhi_num > xinzhis:
            fen /= (1-(xinzhis - limit_xinzhi_num) * 0.5)        # 不满心智毕业数目限制扣分
    return fen


def gene_getvalues(gene):
    values = [0, 0, 0, 1.5, 1]  # 省物资0，省魔方1，缺心智2，金：彩3，彩：装备4
    values[0] += 0.5 * gene[0]
    values[1] += 0.5 * gene[1]
    values[2] += 0.1 * gene[2]
    values[3] += 0.5 * gene[3]
    values[4] += 2 * gene[4]
    return values


def gene_shiyingdu(gene, maindata):
    values1 = gene_getvalues(gene.gene1)
    values2 = gene_getvalues(gene.gene2)
    values3 = gene_getvalues(gene.gene3)
    chan1 = shouyi(values1, maindata, 0)
    chan2 = shouyi(values2, maindata, 1)
    chan3 = shouyi(values3, maindata, 2)
    chan = [chan1, chan2, chan3]
    fen = shiyingdu(maindata, chan)
    return fen


def print_in_excel(gene, maindata):
    values1 = gene_getvalues(gene.gene1)
    values2 = gene_getvalues(gene.gene2)
    values3 = gene_getvalues(gene.gene3)
    chan1 = shouyi(values1, maindata, 0)
    chan2 = shouyi(values2, maindata, 1)
    chan3 = shouyi(values3, maindata, 2)
    chan = [chan1, chan2, chan3]
    day1, day2, zhuangs, xinzhis, wuzi_day, mofang_day = chanchulist(chan)
    celue = chan1[8]
    workbook = xlwt.Workbook(encoding='utf-8')  # 产生工作表，保存数据
    worksheet = workbook.add_sheet("策略表")
    worksheet.write(9, 0, label="金船满破阶段")
    for i in range(30):
        p = i // 10
        p *= 3
        q = i % 10 + 10
        r = i + 1
        if i == chan1[7]:
            worksheet.write(q, p, label=r)
            worksheet.write(q, p + 1, label="刷新")
            continue
        elif i >= chan1[7]:
            j = i - 1
            j = celue[j]
        else:
            j = celue[i]
        string_1 = maindata[j + 2][0]
        worksheet.write(q, p, label=r)
        worksheet.write(q, p + 1, label=string_1)

    worksheet.write(21, 0, label="彩船满破阶段")
    celue = chan2[8]
    for i in range(30):
        p = i // 10
        p *= 3
        q = i % 10 + 22
        r = i + 1
        if i == chan2[7]:
            worksheet.write(q, p, label=r)
            worksheet.write(q, p + 1, label="刷新")
            continue
        elif i >= chan2[7]:
            j = i - 1
            j = celue[j]
        else:
            j = celue[i]
        string_1 = maindata[j + 2][0]
        worksheet.write(q, p, label=r)
        worksheet.write(q, p + 1, label=string_1)

    worksheet.write(33, 0, label="做彩装备阶段")
    celue = chan3[8]
    for i in range(30):
        p = i // 10
        p *= 3
        q = i % 10 + 34
        r = i + 1
        if i == chan3[7]:
            worksheet.write(q, p, label=r)
            worksheet.write(q, p + 1, label="刷新")
            continue
        elif i >= chan3[7]:
            j = i - 1
            j = celue[j]
        else:
            j = celue[i]
        string_1 = maindata[j + 2][0]
        worksheet.write(q, p, label=r)
        worksheet.write(q, p + 1, label=string_1)
    wuzi = str(int(wuzi_day * 0.0365)) + "万"
    mofang = str(int(mofang_day * 365)) + "个"
    day1_str = str(int(day1)) + "天"
    day2_str = str(int(day2)) + "天"

    i = 0
    worksheet.write(i, 0, label="三个阶段综合情况")
    i += 1
    worksheet.write(i, 0, label="年总物资消耗")
    worksheet.write(i, 2, label=wuzi)
    i += 1
    worksheet.write(i, 0, label="年总魔方消耗")
    worksheet.write(i, 2, label=mofang)
    i += 1
    worksheet.write(i, 0, label="金船毕业时间")
    worksheet.write(i, 2, label=day1_str)
    i += 1
    worksheet.write(i, 0, label="彩船毕业时间")
    worksheet.write(i, 2, label=day2_str)
    i += 1
    worksheet.write(i, 0, label="年总彩装产出")
    worksheet.write(i, 2, label=zhuangs)
    i += 1
    worksheet.write(i, 0, label="年总心智产出")
    worksheet.write(i, 2, label=xinzhis)

    workbook.save('策略输出表.xls')



def tuihuo(maindata):
    tuihuo_xunhuan_num = int(maindata[46][1])
    gene_base = Gene(maindata)
    shiying_base = gene_shiyingdu(gene_base, maindata)
    for i in range(200):
        gene_new = Gene(maindata)
        gene_new.newrand()
        fen = gene_shiyingdu(gene_new, maindata)
        if fen > shiying_base:
            shiying_base = fen
            gene_base = copy.deepcopy(gene_new)

    for i in range(tuihuo_xunhuan_num):
        gene_now = copy.deepcopy(gene_base)
        for j in range(5):
            gene = copy.deepcopy(gene_now)
            gene.change()
            fen = gene_shiyingdu(gene, maindata)
            if fen > shiying_base:
                gene_now = copy.deepcopy(gene)
                gene_base = copy.deepcopy(gene)
                shiying_base = gene_shiyingdu(gene_base, maindata)
            else:
                p = (5 - j) * 10
                p = random.randint(0, p)
                if (fen + p) > shiying_base:
                    gene_now = copy.deepcopy(gene)
        if i % 20 == 19:
            str1 = "模拟退火第" + str(i + 1) + "代"
            print(str1, "适应度", shiying_base)
    print_in_excel(gene_base, maindata)
    return 0



fname = "策略限制表.xls"
sname = "Sheet1"
maindata = excel(fname, sname)
values = [0, 160, 0, 2.4, 14.5, 0]    # 物资0，魔方1，咸鱼2，彩金兑换3，彩装兑换4，心智5
way = maindata[0][10]
tuihuo(maindata)
while 1:
    print("按回车关闭程序")
    a = input()
    print(a)
    if a == '':
        break



