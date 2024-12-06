import os.path
from datetime import datetime

import pandas as pd
from numpy import ndarray
from ortools.sat.python import cp_model


class Teacher(object):

    def __init__(self, arr: ndarray[7]):
        self.no = arr[0]
        self.name = arr[1]
        self.s_y = arr[2]
        self.s_n = arr[3]
        self.r_y = arr[4]
        self.r_n = arr[5]
        self.times_limit = int(arr[6])

    def __str__(self):
        return f'T({self.no} {self.name} {self.times_limit})'

    def __repr__(self):
        return self.__str__()


class Subject(object):

    def __init__(self, arr: ndarray[4]):
        self.code = arr[0]
        self.name = arr[1]
        self.date = arr[2].date().strftime('%Y%m%d')
        self.time = arr[3]

    @property
    def apm(self):
        """
        判断下当前时间是 AM、PM
        :return:
        """
        if '-' in self.time:
            stime, etime = self.time.split('-')
        else:
            stime, etime = self.time.split('—')
        if int(etime.split(':')[0]) < 13:
            return 'AM'
        if int(stime.split(':')[0]) > 13:
            return 'PM'
        raise Exception('time range error')

    def __str__(self):
        return f'S({self.code} {self.name} {self.date} {self.time})'

    def __repr__(self):
        return self.__str__()


class Room(object):
    serials = []

    def __init__(self, arr: ndarray[10]):
        self.name = arr[0]
        self.nums = arr[1:]

    def __str__(self):
        return f'R({self.name} {self.nums})'

    def __repr__(self):
        return self.__str__()


def read_data() -> (list[Teacher], list[Subject], list[Room]):
    xlsx_file = os.path.join('in_files', '监考安排.xlsx')
    sheets = pd.read_excel(xlsx_file, sheet_name=None)
    pd_teachers = sheets['监考员设置']
    pd_subjects = sheets['考试科目设置']
    pd_rooms = sheets['考场设置']
    # 构建模型数据
    teachers = [Teacher(d) for d in pd_teachers.values if isinstance(d[1], str)]
    subjects = [Subject(d) for d in pd_subjects.values]
    Room.serials = pd_rooms.columns[1:].tolist()
    rooms = [Room(d) for d in pd_rooms.values]
    #
    return teachers, subjects, rooms


def write_data(res: list[(str, str, str)], teachers: list[Teacher], subjects: list[Subject], rooms: list[Room]):
    if not res:
        return
    # 1 考场
    room_names = [r.name for r in rooms]
    room_data = {
        '考场': room_names
    }
    for s in subjects:
        room_data[s.name] = [[] for _ in rooms]
    for ss in subjects:
        for rs in rooms:
            for s, r, t in res:
                if ss.name == s and rs.name == r:
                    idx = room_names.index(r)
                    room_data[s][idx].append(t)
    # 整理下数据
    for ss in subjects:
        ds = room_data[ss.name]
        for idx in range(len(ds)):
            ds[idx] = ', '.join(ds[idx])
    # 2 老师
    teacher_names = [t.name for t in teachers]
    teacher_data = {
        '老师': teacher_names
    }
    for s in subjects:
        teacher_data[s.name] = [None for _ in teachers]
    for ss in subjects:
        for ts in teachers:
            for s, r, t in res:
                if ss.name == s and ts.name == t:
                    idx = teacher_names.index(t)
                    teacher_data[s][idx] = r
    # 写数据
    date_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_file = os.path.join('out_files', f'监考安排_{date_str}.xlsx')
    with pd.ExcelWriter(out_file) as writer:
        pd.DataFrame(room_data).to_excel(excel_writer=writer, sheet_name='考场安排')
        pd.DataFrame(teacher_data).to_excel(excel_writer=writer, sheet_name='监考员安排')


def solve(teachers: list[Teacher], subjects: list[Subject], rooms: list[Room]):
    # 1. 初始化模型
    model = cp_model.CpModel()

    # 2. 创建变量
    inv_schedule = {}
    teacher_date = {}
    for s in subjects:
        for r in rooms:
            for t in teachers:
                var_name = f'subject_{s.name}_room_{r.name}_teacher_{t.name}'
                inv_schedule[(s.name, r.name, t.name)] = model.new_bool_var(var_name)
        for t in teachers:
            var_name = f'teacher_{t.name}_date_{s.date}'
            teacher_date[(t.name, s.date, s.apm)] = model.new_bool_var(var_name)

    # 2. 添加约束
    # 同一科目，同一个老师，最多只能出现一次
    for s in subjects:
        for t in teachers:
            model.add_at_most_one(inv_schedule[(s.name, r.name, t.name)] for r in rooms)
    # 同一科目 某教室，只能出现指定数量老师
    for s in subjects:
        for r in rooms:
            idx = r.serials.index(s.name)
            nums = r.nums[idx]
            model.add(sum(inv_schedule[(s.name, r.name, t.name)] for t in teachers) == nums)
    # 同一老师限制最长场次
    for t in teachers:
        model.add(sum(inv_schedule[(s.name, r.name, t.name)] for s in subjects for r in rooms) <= t.times_limit)
    # 限制 必监考科目/不监考科目
    for t in teachers:
        # 必监考科目
        if isinstance(t.s_y, str):
            sys = t.s_y.split('/')
            for sy in sys:
                model.add_at_least_one(inv_schedule[(sy, r.name, t.name)] for r in rooms)
        # 不监考科目
        if isinstance(t.s_n, str):
            sns = t.s_n.split('/')
            for sn in sns:
                model.add(sum(inv_schedule[(sn, r.name, t.name)] for r in rooms) == 0)
    # 限制 必监考教室/不监考教室
    for t in teachers:
        # 必监考教室
        if isinstance(t.r_y, str):
            rys = t.r_y.split('/')
            for ry in rys:
                # 逻辑1: 这个老师至少在这个教室一次
                # model.add_at_least_one(inv_schedule[(s.name, ry, t.name)] for s in subjects)
                # 逻辑2: 这个老师只能在这个教室
                model.add(
                    sum(inv_schedule[(s.name, r.name, t.name)] for s in subjects for r in rooms if r.name != ry) == 0)
        # 不监考教室
        if isinstance(t.r_n, str):
            rns = t.r_n.split('/')
            for rn in rns:
                model.add(sum(inv_schedule[(s.name, rn, t.name)] for s in subjects) == 0)
    # 工作时间
    for t in teachers:
        for s in subjects:
            model.add_bool_or([inv_schedule[(s.name, r.name, t.name)].Not() for r in rooms]).only_enforce_if(
                teacher_date[(t.name, s.date, s.apm)].Not())
            model.add_bool_or([inv_schedule[(s.name, r.name, t.name)] for r in rooms]).only_enforce_if(
                teacher_date[(t.name, s.date, s.apm)])

    # 4. 定义目标函数
    # 尽量不要分散排 -> 每个老师工作的日期(date+AM/PM)数量最少 -> 全部老师工作日期最少
    model.minimize(sum(teacher_date[(t.name, s.date, s.apm)] for s in subjects for t in teachers))

    # 5. 添加求解器
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    # 6. 处理结果
    result = []
    if status == cp_model.OPTIMAL:
        for s in subjects:
            for r in rooms:
                for t in teachers:
                    if solver.boolean_value(inv_schedule[(s.name, r.name, t.name)]):
                        print(f'科目: {s.name}, 教室: {r.name}, 老师: {t.name}')
                        result.append((s.name, r.name, t.name))

    else:
        print('no solution')

    # 7. 返回结果
    return result


def main():
    teachers, subjects, rooms = read_data()
    result = solve(teachers, subjects, rooms)
    write_data(result, teachers, subjects, rooms)


if __name__ == '__main__':
    main()
