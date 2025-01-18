import os
import os.path as osp
import time

from utils import parse_excel, parse_songs, make_period

plan_excel = "跑步计划.xlsx"
dir_songs = r"D:\CloudMusic"
dir_out = "outputs"
os.makedirs(dir_out, exist_ok=True)

run_songs, walk_songs = parse_songs(dir_songs)
data = parse_excel(plan_excel)


def handle_one_training(item):
    t0 = time.time()
    loop = item.get("loop_times")
    run_mins = item.get("run_mins")
    walk_mins = item.get("walk_mins")
    total_time = item.get("total_mins")
    process_id = f"{item.get('week')}-第{item.get('idx_week')}次训练"

    periods = []
    for i, (run_m, walk_m) in enumerate(zip(run_mins, walk_mins)):
        first_run_time = run_m if i == 0 else None
        periods.append(make_period(run_songs, run_m, run_flag=False, nxt_time=walk_mins[i], process_id=process_id,
                                   first_run_time=first_run_time))
        if walk_mins[i] == 0:
            continue
        nxt_time = run_mins[i + 1] if i < loop - 1 else None
        periods.append(make_period(walk_songs, walk_m, run_flag=True, nxt_time=nxt_time, process_id=process_id))
    merge = sum(periods)
    out_song = osp.join(dir_out, f"{process_id}-总时长{total_time}分钟.mp3")
    merge.export(out_song, format="mp3")
    print(f"{out_song}制作完成，耗时{(time.time() - t0) / 60:.2f} 分钟")
    return True


def main():
    t0 = time.time()

    for item in data:
        handle_one_training(item)
        break

    # with futures.ProcessPoolExecutor(max_workers=1) as executor:
    #     executor.map(handle_one_training, data)

    print(f"所有任务完成，耗时{(time.time() - t0) / 60 :.2f}分钟")


if __name__ == '__main__':
    main()
