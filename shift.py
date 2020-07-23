# -*- coding: utf-8 -*-
import random
from scoop import futures
import pandas as pd
from deap import base
from deap import creator
from deap import tools
from deap import cma

# 従業員を表すクラス
# 0が休み 1が通常、2が待機入り、3が夜勤、4が夜勤明け

def read_excel():
    num_people_per_day = []
    holidays = []
    indexes = []
    columns_for_excel = []
    # Excelの読み込み
    df = pd.read_excel("hoge.xlsx")
    # dfの欲しい場所だけを抜き取る
    index = df.iloc[4:27, 0:1]
    index = index.reset_index(drop=True)
    index.columns = [i + 1 for i in range(len(index.columns))]
    for column in range(len(index.columns)):
        for row in range(len(index)):
            indexes.append(index[column + 1][row])
    column_for_excel = df.iloc[2:3, 1:]
    column_for_excel = column_for_excel.reset_index(drop=True)
    column_for_excel.columns = [i + 1 for i in range(len(column_for_excel.columns))]
    for hoge in range(len(column_for_excel.columns)):
        for wow in range(len(column_for_excel)):
            columns_for_excel.append(column_for_excel[hoge + 1][wow])
    df = df.iloc[4:, 1:]
    # エラー値を0に置換
    df = df.fillna(5)
    # 希望休（◎）を0に置換
    df = df.replace("◎", 0)
    # 休日数のテーブルを取得
    days = df.iloc[24:, :].reset_index(drop=True)
    days.columns = [i + 1 for i in range(len(days.columns))]
    for columns in range(len(days.columns)):
        for rows in range(len(days)):
            num_people_per_day.append(days[columns + 1][rows])
    # 基礎テーブルを取得
    holiday = df.iloc[:23, :].reset_index(drop=True)
    holiday.columns = [i + 1 for i in range(len(holiday.columns))]
    for columns in range(len(holiday.columns)):
        for rows in range(len(holiday)):
            holidays.append(holiday[columns + 1][rows])
    return holidays, num_people_per_day, indexes, columns_for_excel

# シフトを表すクラス
# 内部的には (5 * 21日 + 2 * 5日) * 23人 = 2645次元のタプルで構成される
# 内部的には (31日) * 23人 = 713次元のタプルで構成される


class Shift(object):

    SHIFT_BOXES = ['sat', 'sun', 'mon', 'tue', 'wed', 'thu', 'fri']

    holidays, num_people_per_day, indexes, columns = read_excel()

    def __init__(self, list):
        if list is None:
            self.make_sample()
        else:
            self.list = list
        # self.employees = []

    # ランダムなデータを生成
    def make_sample(self):
        sample_list = []
        for num in range(713):
            sample_list.append(random.randint(0, 4))
        self.list = tuple(sample_list)

    # タプルを1ユーザ単位に分割
    def slice(self):
        sliced = []
        start = 0
        for num in range(23):
            sliced.append(self.list[start:(start + 31)])
            start = start + 31
        # print(sliced)
        return tuple(sliced)

    # TSV形式でアサイン結果の出力をする
    # columnとindexのデータを取得して代入
    def print_tsv(self):
        result = []
        for line in self.slice():
            for i, num in enumerate(line):
                if num == 0:
                    line[i] = "×"
                elif num == 1:
                    line[i] = ""
                elif num == 2:
                    line[i] = "★"
                elif num == 3:
                    line[i] = "▲"
                elif num == 4:
                    line[i] = "○"
            result.append(line)

        df = pd.DataFrame(result,
                          index=self.indexes, columns=self.columns)
        with pd.ExcelWriter('happy.xlsx') as writer:
            df.to_excel(writer, sheet_name='sheet1')

    def holidays_index(self):
        return [i for i, x in enumerate(self.holidays) if x == 0]

    # ユーザ番号を指定してコマ名を取得する
    def get_boxes_by_user(self, user_no):
        line = self.slice()[user_no]
        return self.line_to_box(line)

    def get_boxes_by_user2(self, user_no):
        line = self.slice()[user_no]
        return self.line_to_box2(line)

    def get_boxes_by_user3(self, user_no):
        line = self.slice()[user_no]
        return self.line_to_box3(line)

    def get_boxes_by_user4(self, user_no):
        line = self.slice()[user_no]
        return self.line_to_box4(line)

    def line_to_box2(self, line):
        result = []
        index = 0
        for e in line:
            if e == 2 or e == 3:
                result.append(self.SHIFT_BOXES[index])
            index += 1
            if index == 7:
                result.append("mark")
                index = 0
        return result

    def line_to_box3(self, line):
        result = []
        index = 0
        for e in line:
            if e == 1 or e == 3:
                result.append(self.SHIFT_BOXES[index])
            index += 1
            if index == 7:
                index = 0
        return result

    def line_to_box4(self, line):
        result = []
        index = 0
        for e in line:
            if e == 2:
                result.append(self.SHIFT_BOXES[index])
            index += 1
            if index == 7:
                index = 0
        return result

    # 1ユーザ分のタプルからコマ名を取得する
    def line_to_box(self, line):
        result = []
        index = 0
        for e in line:
            if e == 2:
                result.append(self.SHIFT_BOXES[index])
            index += 1
            if index == 7:
                index = 0
        return result

    # コマ番号を指定してアサインされているユーザ番号リストを取得する
    def get_user_nos_by_box_index(self, box_index):
        user_nos = []
        index = 0
        for line in self.slice():
            if line[box_index] == 1:
                user_nos.append(index)
            index += 1
        return user_nos

    def get_user_nos_by_box_index2(self, box_index, num):
        user_nos = []
        index = 0
        for i, line in enumerate(self.slice()):
            if line[box_index] == num:
                user_nos.append(index)
            index += 1
        return user_nos

    # コマ名を指定してアサインされているユーザ番号リストを取得する
    def get_user_nos_by_box_name(self, box_name):
        box_index = self.SHIFT_BOXES.index(box_name)
        return self.get_user_nos_by_box_index(box_index)

    # 想定人数と実際の人数の差分を取得する
    def abs_people_between_need_and_actual(self):
        result = []
        index = 0
        num = 0
        for need in self.num_people_per_day:
            actual = len(self.get_user_nos_by_box_index2(index, num))
            result.append(abs(need - actual))
            num += 1
            if num == 5:
                num = 0
                index += 1
        return result

    # 平日待機・遅番待機・夜勤は1週間で１〜2回
    # ここから
    # 1週間をどう設定するか、2,3で既に待機・夜勤がある日は取得できている
    def few_box_per_week(self):
        counter = 0
        for user_no in range(23):
            boxes = self.get_boxes_by_user2(user_no)
            waiting_counter = []
            for box in boxes:
                if box == "mark":
                    num_of_tar = len(waiting_counter)
                    if num_of_tar >= 3:
                        counter += 1
                    waiting_counter = []
                else:
                    waiting_counter.append(box)
        return counter

    # 土日待機、金曜待機は月1回
    def one_per_month_for_on_call(self):
        counter = 0
        for user_no in range(23):
            boxes = self.get_boxes_by_user(user_no)
            if boxes.count('sat') > 1:
                counter += 1
            if boxes.count('san'):
                counter += 1
            if boxes.count('fri') > 1:
                counter += 1
        return counter

    def request_holidays(self):
        counter = 0
        for i in self.holidays_index():
            if self.list[i] != 0:
                counter += 1
        return counter

    # 週末は0か2か4
    def two_or_zero_weekend(self):
        result = []
        for user_no in range(23):
            boxes = self.get_boxes_by_user3(user_no)
            for box in boxes:
                if box == 'san' or box == 'sat':
                    result.append(user_no)
        return result

    def night_shift(self):
        counter = 0
        for user_no in range(23):
            boxes = self.slice()[user_no]
            for i, box in enumerate(boxes):
                if box == 3:
                    if i + 1 < len(boxes):
                        if boxes[i+1] != 4:
                            counter += 1
                    if i - 1 >= 0:
                        if boxes[i-1] != 0 and boxes[i-1] != 1:
                            counter += 1
        return counter

    def weekend_num(self):
        counter = 0
        for index in range(31):
            actual = len(self.get_user_nos_by_box_index2(index, 2))
            if index % 7 == 1 or index % 7 == 0:
                if actual != 4:
                    counter += 1
            else:
                if actual != 3:
                    counter += 1
        return counter

    def night_shift_num(self):
        counter = 0
        for need in range(31):
            if need % 7 != 1 and need % 7 != 0:
                actual = len(self.get_user_nos_by_box_index2(need, 3))
                if actual != 1:
                    counter += 1
        return counter

# 従業員定義
creator.create("FitnessPeopleCount", base.Fitness, weights=(-1.0,-1.0, -1.0, -1.0, -1.0, -1.0, -10.0, -1.0))
creator.create("Individual", list, fitness=creator.FitnessPeopleCount)

toolbox = base.Toolbox()

toolbox.register("map", futures.map)


def random_man():
    dice = [1, 2]
    w = [4, 1]
    return random.choices(dice, w)[0]

toolbox.register("attr_bool", random_man)
toolbox.register("individual", tools.initRepeat, creator.Individual, toolbox.attr_bool, 713)
toolbox.register("population", tools.initRepeat, list, toolbox.individual)

# 713 % 31 == 2, 713 % 31 == 3, 713 % 31 == 4, 713 % 5 == 5, ...週一で夜勤を入れる
def night_shift_one_per_day(individual):
    result = []
    s = Shift(individual)
    li = []
    for box in s.list:
        for j in [2, 3, 4, 5, 6, 9, 10, 11, 12, 13, 16, 17, 18, 19, 20, 23, 24, 25, 26, 27, 30]:
            for i, num in enumerate(box):
                if i % 31 == j and num == 1:
                    li.append(i)
            ran_num = random.choice(li)
            box[ran_num] = 3
            li.clear()
        result.append(box)
    return result

# 夜勤の次の日は夜勤明け,夜勤の前は日勤待機・遅番待機にならない
def night_per_week(individual):
    counter = [30, 60, 90, 120, 150, 180, 210, 240, 270, 300, 330,
               360, 390, 420, 450, 480, 510, 540, 570, 600, 630, 660, 690]
    result = []
    s = Shift(individual)
    for box in s.list:
        for i, num in enumerate(box):
            if num == 3 and i not in counter:
                if i+1 < len(box):
                    box[i + 1] = 4
                if i-1 >= 0:
                    box[i - 1] = random.randint(0, 1)
        result.append(box)
    return result


# 週末は0か2か4
def weekend(individual):
    counter = 0
    result = []
    s = Shift(individual)
    for box in s.list:
        for i, num in enumerate(box):
            if num == 1 or num == 3 or num == 2:
                if counter % 7 == 0:
                    if i + 1 < len(box):
                        ran_num = random.choices([0, 2], [2, 1])[0]
                        # ran_num = random.choice([0, 2])
                        box[i] = ran_num
                        box[i + 1] = ran_num
                elif counter % 7 == 1:
                    if i - 1 >= 0:
                        ran_num = random.choices([0, 2], [1.7, 1])[0]
                        # ran_num = random.choice([0, 2])
                        box[i] = ran_num
                        box[i - 1] = ran_num
            counter += 1
            if counter == 31:
                counter = 0
        result.append(box)
    return result


def evalShift(individual):
    # print(individual)
    s = Shift(individual)
    # s.employees = employees

    # 想定人数とアサイン人数の差
    people_count_sub_sum = sum(s.abs_people_between_need_and_actual()) / 713.0

    few_box_per_week = s.few_box_per_week() / 713.0

    one_per_month_for_on_call = s.one_per_month_for_on_call() / 713.0

    holiday = s.request_holidays() / 713.0

    two_or_zero_weekend = len(s.two_or_zero_weekend()) / 23.0

    night_shift = s.night_shift() / 713.0

    weekend_num = s.weekend_num() / 31.0

    night_shift_num = s.night_shift_num() / 21.0

    return people_count_sub_sum, few_box_per_week, one_per_month_for_on_call,\
           holiday, two_or_zero_weekend, night_shift, weekend_num, night_shift_num


def request_holiday(individual):
    result = []
    s = Shift(individual)
    for box in s.list:
        holidays_list = s.holidays_index()
        for days in holidays_list:
            box[days] = 0
        result.append(box)
    return result

toolbox.register("evaluate", evalShift)
toolbox.register("night_shift_one_per_day", night_shift_one_per_day)
toolbox.register("night_per_week", night_per_week)
toolbox.register("weekend", weekend)
toolbox.register("request_holiday", request_holiday)
# 交叉関数を定義(二点交叉)
toolbox.register("mate", tools.cxTwoPoint)

# 変異関数を定義(ビット反転、変異隔離が5%ということ?)
toolbox.register("mutate", tools.mutFlipBit, indpb=0.05)

# 選択関数を定義(トーナメント選択、tournsizeはトーナメントの数？)
toolbox.register("select", tools.selTournament, tournsize=3)

if __name__ == '__main__':
    # 初期集団を生成する
    pop = toolbox.population(n=300)
    pop = toolbox.night_shift_one_per_day(pop)
    pop = toolbox.night_per_week(pop)
    pop = toolbox.weekend(pop)
    pop = toolbox.request_holiday(pop)
    CXPB, MUTPB, NGEN = 0.6, 0.5, 500  # 交差確率、突然変異確率、進化計算のループ回数

    print("進化開始")

    # 初期集団の個体を評価する
    fitnesses = list(map(toolbox.evaluate, pop))
    for ind, fit in zip(pop, fitnesses):  # zipは複数変数の同時ループ
        # 適合性をセットする
        ind.fitness.values = fit

    # print("  %i の個体を評価" % len(pop))

     # 進化計算開始
    for g in range(NGEN):
        # print("-- %i 世代 --" % g)

        # 選択
        # 次世代の個体群を選択
        offspring = toolbox.select(pop, len(pop))
        # 個体群のクローンを生成
        offspring = list(map(toolbox.clone, offspring))

        # 選択した個体群に交差と突然変異を適応する

        # 交叉
        # 偶数番目と奇数番目の個体を取り出して交差
        for child1, child2 in zip(offspring[::2], offspring[1::2]):
            if random.random() < CXPB:
                toolbox.mate(child1, child2)
                # 交叉された個体の適合度を削除する
                del child1.fitness.values
                del child2.fitness.values

        # 変異
        for mutant in offspring:
            if random.random() < MUTPB:
                toolbox.mutate(mutant)
                del mutant.fitness.values

        # 適合度が計算されていない個体を集めて適合度を計算
        invalid_ind = [ind for ind in offspring if not ind.fitness.valid]
        fitnesses = map(toolbox.evaluate, invalid_ind)
        for ind, fit in zip(invalid_ind, fitnesses):
            ind.fitness.values = fit
        # 次世代群をoffspringにする
        pop[:] = offspring

        # すべての個体の適合度を配列にする
        index = 1
        for v in ind.fitness.values:
          fits = [v for ind in pop]

          length = len(pop)
          mean = sum(fits) / length
          sum2 = sum(x*x for x in fits)
          std = abs(sum2 / length - mean**2)**0.5

          index += 1

    print("-- 進化終了 --")

    best_ind = tools.selBest(pop, 1)[0]
    print("最も優れていた個体: %s, %s" % (best_ind, best_ind.fitness.values))
    s = Shift(best_ind)
    s.print_tsv()
