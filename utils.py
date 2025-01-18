import os
import os.path as osp
import random
from glob import glob
from uuid import uuid1

import pyttsx3
import xlrd
from pydub import AudioSegment

ENGINE = pyttsx3.init()
SEC = 1000
dir_temp = "temp"
os.makedirs(dir_temp, exist_ok=True)


def make_data(v, loop_v):
    if isinstance(v, str):
        items = v.strip().split(' ')
        return list(map(int, items))
    else:
        if not isinstance(loop_v, float) and not isinstance(loop_v, int):
            raise Exception(f"循环次数应该是一个整数，但是它现在是{loop_v}")
        return [int(v)] * int(loop_v)


def parse_excel(file):
    wb = xlrd.open_workbook(filename=file)
    sheet = wb.sheet_by_index(0)
    merged = sheet.merged_cells
    merged.sort(key=lambda x: x[0])
    # print(merged)
    res = []
    for (rlow, rhigh, clow, chigh) in merged:
        if not (clow == 0 and chigh == 1):
            continue
        for idx_week, idx_row in enumerate(range(rlow, rhigh)):
            run_v = sheet.cell(rlow + idx_week, clow + 2).value
            walk_v = sheet.cell(rlow + idx_week, clow + 3).value
            loop_v = sheet.cell(rlow + idx_week, clow + 4).value
            run_mins = make_data(run_v, loop_v)
            walk_mins = make_data(walk_v, loop_v)
            assert len(run_mins) == len(walk_mins)
            total_mins = sum(run_mins) + sum(walk_mins)
            res.append({
                "week": sheet.cell(rlow, clow).value,
                "idx_week": idx_week + 1,
                "run_mins": run_mins,
                "walk_mins": walk_mins,
                "loop_times": len(run_mins),
                "total_mins": total_mins
            })
    return res


def parse_songs(dir_songs):
    dir_run = osp.join(dir_songs, "run")
    run = glob(osp.join(dir_run, "*.mp3")) + glob(osp.join(dir_run, "*.flac"))
    dir_walk = osp.join(dir_songs, "walk")
    walk = glob(osp.join(dir_walk, "*.mp3")) + glob(osp.join(dir_walk, "*.flac"))
    return run, walk


def concat_song(songs, period_time):
    accu = 0
    shuffle_songs = []
    merge = []
    while accu < period_time:
        if not shuffle_songs:
            random.shuffle(songs)
            shuffle_songs = songs.copy()
        mp3_fn = shuffle_songs.pop()
        # print(mp3_fn)
        song = AudioSegment.from_file(mp3_fn)
        merge.append(song)
        accu += song.duration_seconds * SEC
    return sum(merge)


def make_period(song_fns, period_time, run_flag, nxt_time, process_id, first_run_time=None):
    period_time *= 60 * SEC
    period_time -= 2 * SEC
    song = concat_song(song_fns, period_time)
    song_left = song[:period_time]
    period = song_left.fade_out(10 * SEC)

    # 为第一段运动添加提示：
    if first_run_time:
        tp_mp3 = make_hint(run_flag, nxt_time, process_id, first_run_time=first_run_time)
        hint = AudioSegment.from_file(tp_mp3) + 12
        period = hint + period
        os.remove(tp_mp3)

    tmp_mp3 = make_hint(run_flag, nxt_time, process_id, complete=False if nxt_time else True)
    hint = AudioSegment.from_file(tmp_mp3) + 12
    period = period + hint
    os.remove(tmp_mp3)
    return period


def make_hint(run_flag, time, process_id, complete=False, first_run_time=None):
    if not complete:
        msg = f"下一个动作，{'跑步' if run_flag else '走路'}{time}分钟"
    else:
        msg = "恭喜你完成本次训练，休息一下吧"
    if first_run_time:
        msg = f"第一个动作，跑步{first_run_time}分钟"
    res = osp.join(dir_temp, f"{process_id}-{uuid1()}.mp3")
    ENGINE.save_to_file(msg, res)
    ENGINE.runAndWait()
    return res


if __name__ == '__main__':
    dir_songs = r"D:\CloudMusic"
    res = parse_songs(dir_songs)
    for r in res:
        print(r)

    # run_plan_excel = "跑步计划.xlsx"
    # res = parse_excel(run_plan_excel)
    # for r in res:
    #     print(r)
    #
