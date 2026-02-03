# -*- coding: utf-8 -*-
"""
实验名称: 短视频脑电情绪实验 (V7.2 数据规范化版)
修改内容:
  1. [CSV优化] 修正了 CSV 文件头部的列名，使其与录入信息完全对应且专业化（如“习惯”->“每日刷短视频时长”）。
  2. [UI文案] 录入界面的标签更新为更严谨的学术用语。
  3. [Core] 保持 OpenCV + Winsound 强同步内核 + 强退功能。
"""

from psychopy import visual, core, event, gui
import os
import cv2
import winsound
import time
import math
import re
import ctypes 

# ================= 1. 全局配置 =================
IS_DEBUG = False # False = 正式模式

VIDEO_DIR = r"E:\studyResource\同步文件夹\exp02\Videos" 
VIDEO_FILES = [f"{i}.mp4" for i in range(1, 8)]
DATA_DIR = "data-2"
TIME_REST = 1 if IS_DEBUG else 10

# UI 配色 (用于实验界面)
COLOR_BG = '#121212'           
COLOR_TEXT_MAIN = '#FFFFFF'    
COLOR_ACCENT = '#00E5FF'       
FONT_MAIN = 'Microsoft YaHei'      

try:
    from moviepy.editor import VideoFileClip
    HAS_MOVIEPY = True
except ImportError:
    try:
        from moviepy.video.io.VideoFileClip import VideoFileClip
        HAS_MOVIEPY = True
    except ImportError:
        HAS_MOVIEPY = False

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# ================= 2. 核心功能：强制退出检测 =================
def check_force_quit():
    """Shift + Ctrl + Alt 强制退出"""
    shift = ctypes.windll.user32.GetKeyState(0x10) & 0x8000
    ctrl  = ctypes.windll.user32.GetKeyState(0x11) & 0x8000
    alt   = ctypes.windll.user32.GetKeyState(0x12) & 0x8000
    if shift and ctrl and alt:
        try: cv2.destroyAllWindows()
        except: pass
        if 'win' in globals(): win.close()
        core.quit()

# ================= 3. 信息录入 (文案标准化) =================
subject_info = {}

def get_subject_info_traditional():
    """
    使用原生弹窗，文案已更新为标准学术用语
    """
    # 记忆上次输入
    last_id = ''
    last_age = ''
    last_gender = '男'
    last_duration = '>120min'
    last_impact = '中'

    while True:
        check_force_quit()
        
        # 定义标准的学术化标签
        LABEL_ID = '受试者编号 (3位数字)'
        LABEL_AGE = '年龄 (0-99)'
        LABEL_GENDER = '性别'
        LABEL_DURATION = '每日刷短视频时长' # 更正为全称
        LABEL_IMPACT = '短视频影响程度'     # 更正为全称

        info = {
            LABEL_ID: last_id,
            LABEL_AGE: last_age,
            LABEL_GENDER: ['男', '女'],
            LABEL_DURATION: ['>120min', '60-120min', '30-60min', '<30min'],
            LABEL_IMPACT: ['小', '中', '大']
        }
        
        # 回填逻辑
        if last_gender in info[LABEL_GENDER]:
            info[LABEL_GENDER].remove(last_gender)
            info[LABEL_GENDER].insert(0, last_gender)
        if last_duration in info[LABEL_DURATION]:
            info[LABEL_DURATION].remove(last_duration)
            info[LABEL_DURATION].insert(0, last_duration)
        if last_impact in info[LABEL_IMPACT]:
            info[LABEL_IMPACT].remove(last_impact)
            info[LABEL_IMPACT].insert(0, last_impact)

        dlg = gui.DlgFromDict(dictionary=info, title='受试者基本信息录入', 
                              sortKeys=False, 
                              order=[LABEL_ID, LABEL_AGE, LABEL_GENDER, LABEL_DURATION, LABEL_IMPACT])
        
        if not dlg.OK: core.quit()
        
        sid = info[LABEL_ID].strip()
        age = info[LABEL_AGE].strip()
        gender = info[LABEL_GENDER]
        duration = info[LABEL_DURATION]
        impact = info[LABEL_IMPACT]
        
        # 更新记忆
        last_id, last_age, last_gender, last_duration, last_impact = sid, age, gender, duration, impact
        
        # 校验
        error_msg = ""
        if not re.match(r'^\d{3}$', sid):
            error_msg = f"编号 '{sid}' 不合法！\n必须为 3 位数字 (如 001)。"
        elif not age.isdigit() or not (0 <= int(age) <= 99):
            error_msg = f"年龄 '{age}' 不合法！\n请输入 0-99 之间的整数。"
            
        if error_msg:
            err_dlg = gui.Dlg(title="输入错误")
            err_dlg.addText(error_msg)
            err_dlg.show()
        else:
            return {
                'id': sid,
                'age': age,
                'gender': gender,
                'duration': duration,
                'impact': impact
            }

# 运行录入
subject_data = get_subject_info_traditional()
sid_str = subject_data['id']
final_filename = os.path.join(DATA_DIR, f"test{sid_str}-2.csv")

# ================= 4. 数据文件初始化 (标准化表头) =================
if not os.path.exists(final_filename):
    with open(final_filename, 'w', encoding='utf-8-sig') as f:
        # 写入受试者元数据 (Header) - 这里的标签与录入界面完全对应
        f.write("【受试者基本信息】\n")
        f.write(f"受试者编号,{subject_data['id']}\n")
        f.write(f"性别,{subject_data['gender']}\n")
        f.write(f"年龄,{subject_data['age']}\n")
        f.write(f"每日刷短视频时长,{subject_data['duration']}\n") # 已修正为全称
        f.write(f"短视频影响程度,{subject_data['impact']}\n")     # 已修正为全称
        f.write("\n")
        
        # 写入实验数据表头 (Table Header)
        # 保持列名简洁明了，方便 Excel 筛选
        header = ["阶段类型", "视频文件名", "评分/选择", "反应时(ms)"]
        f.write(",".join(header) + "\n")

def log_data(stage, vid_name, choice, rt):
    # 数据记录，确保无多余空格，符合 CSV 规范
    row = [stage, vid_name, choice, str(rt)]
    with open(final_filename, 'a', encoding='utf-8-sig') as f:
        f.write(",".join(row) + "\n")

# ================= 5. 窗口初始化 =================
win = visual.Window(size=[1920, 1080], fullscr=True, monitor="testMonitor", 
                    units="height", color=COLOR_BG, allowGUI=False)
win.mouseVisible = False

# ================= 6. 核心逻辑 (OpenCV + Winsound) =================

def check_audio():
    txt = visual.TextStim(win, text="资源加载中...", font=FONT_MAIN, pos=(0,0), height=0.04, color=COLOR_ACCENT)
    txt.draw(); win.flip()
    check_force_quit()

    if not HAS_MOVIEPY: return
    for vid_file in VIDEO_FILES:
        mp4_path = os.path.join(VIDEO_DIR, vid_file)
        wav_path = os.path.join(VIDEO_DIR, os.path.splitext(vid_file)[0] + ".wav")
        if os.path.exists(mp4_path) and not os.path.exists(wav_path):
            try:
                clip = VideoFileClip(mp4_path)
                if clip.audio: clip.audio.write_audiofile(wav_path, fps=44100, logger=None)
                clip.close()
            except: pass
    check_force_quit()

def play_video(video_path):
    base, _ = os.path.splitext(video_path)
    audio_path = base + ".wav"
    has_audio = os.path.exists(audio_path)

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened(): return

    fps = cap.get(cv2.CAP_PROP_FPS)
    if fps <= 0: fps = 30
    duration = int(cap.get(cv2.CAP_PROP_FRAME_COUNT)) / fps
    frame_interval = 1.0 / fps 

    window_name = 'Video_Player'
    cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)
    cv2.setWindowProperty(window_name, cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
    cv2.setWindowProperty(window_name, cv2.WND_PROP_TOPMOST, 1)

    if has_audio:
        winsound.PlaySound(audio_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
    
    start_time = time.perf_counter()
    current_frame_idx = 0
    win.mouseVisible = False
    
    while cap.isOpened():
        check_force_quit()
        
        elapsed = time.perf_counter() - start_time
        if elapsed > duration + 0.5: break 
        
        target_frame_idx = int(elapsed * fps)
        
        if current_frame_idx < target_frame_idx:
            ret, _ = cap.read()
            current_frame_idx += 1
            if not ret: break
            continue 
        elif current_frame_idx > target_frame_idx:
            wait_time = (current_frame_idx * frame_interval) - elapsed
            if wait_time > 0: time.sleep(wait_time)
        
        ret, frame = cap.read()
        current_frame_idx += 1
        if not ret: break
        
        cv2.imshow(window_name, frame)
        cv2.waitKey(1) 

    winsound.PlaySound(None, winsound.SND_PURGE)
    cap.release()
    cv2.destroyAllWindows()
    win.mouseVisible = False

# ================= 7. 评分与休息 UI =================

def draw_option_btn(text, pos, is_selected):
    color = COLOR_ACCENT if is_selected else '#2A2A2A'
    opacity = 1.0 if is_selected else 0.8
    txt_col = '#121212' if is_selected else '#CCCCCC'
    visual.Rect(win, width=0.6, height=0.09, pos=pos, fillColor=color, opacity=opacity, lineColor=None).draw()
    visual.TextStim(win, text=text, font=FONT_MAIN, pos=pos, height=0.035, color=txt_col, bold=True).draw()

def get_rating(title, options, key_map):
    event.clearEvents()
    start_y = 0.1
    gap_y = 0.14
    selected_idx = -1
    
    while True:
        check_force_quit()

        visual.TextStim(win, text=title, font=FONT_MAIN, pos=(0, 0.35), height=0.05, color='white', bold=True).draw()
        visual.TextStim(win, text="按键盘对应字母选择", font=FONT_MAIN, pos=(0, 0.28), height=0.025, color='#666666').draw()
        
        for i, opt in enumerate(options):
            draw_option_btn(opt, (0, start_y - i*gap_y), (i == selected_idx))
        
        win.flip()
        
        keys = event.getKeys(keyList=list(key_map.keys()) + ['escape'])
        if keys:
            k = keys[0]
            if k == 'escape': win.close(); core.quit()
            selected_idx = key_map[k]
            
            visual.TextStim(win, text=title, font=FONT_MAIN, pos=(0, 0.35), height=0.05, color='white', bold=True).draw()
            visual.TextStim(win, text="按键盘对应字母选择", font=FONT_MAIN, pos=(0, 0.28), height=0.025, color='#666666').draw()
            for i, opt in enumerate(options):
                draw_option_btn(opt, (0, start_y - i*gap_y), (i == selected_idx))
            win.flip()
            core.wait(0.3)
            return k.upper()

def run_rest(sec):
    timer = core.CountdownTimer(sec)
    ring = visual.Circle(win, radius=0.15, edges=128, lineColor=COLOR_ACCENT, lineWidth=3, fillColor=None)
    
    while timer.getTime() > 0:
        check_force_quit()
        
        t = timer.getTime()
        ring.radius = 0.15 + 0.01 * math.sin(t*5)
        visual.TextStim(win, text="休息阶段", font=FONT_MAIN, pos=(0, 0.25), height=0.04, color='white').draw()
        visual.TextStim(win, text=f"{int(t)+1}", font=FONT_MAIN, pos=(0, 0), height=0.1, bold=True, color=COLOR_ACCENT).draw()
        ring.draw()
        win.flip()
        if event.getKeys(['escape']): win.close(); core.quit()

def show_welcome():
    event.clearEvents()
    while True:
        check_force_quit()
        
        visual.TextStim(win, text="短视频脑电认知实验", font=FONT_MAIN, pos=(0, 0.35), 
                    height=0.07, bold=True, color=COLOR_ACCENT).draw()
        
        intro_text = (
            "欢迎你参与本次实验。\n\n"
            "【实验流程】\n"
            "1. 观看 7 段短视频广告，视频结束后进行打分。\n"
            "2. 每次打分后休息 10 秒。\n"
            "3. 全部结束后进行总评。\n\n"
            "注意：按 Esc 可退出 (管理员强退：Shift+Ctrl+Alt)"
        )
        visual.TextStim(win, text=intro_text, font=FONT_MAIN, pos=(0, 0.0), 
                        height=0.03, color=COLOR_TEXT_MAIN, alignText='left', wrapWidth=1.0).draw()
        
        visual.TextStim(win, text="— 按 空格键 (Space) 开始 —", font=FONT_MAIN, 
                        pos=(0, -0.4), height=0.03, color='#666666').draw()
        
        win.flip()
        if event.getKeys(['space']): break
        if event.getKeys(['escape']): win.close(); core.quit()

# ================= 8. 主程序执行 =================
try:
    check_audio()
    show_welcome()

    for vid_file in VIDEO_FILES:
        path = os.path.join(VIDEO_DIR, vid_file)
        if not os.path.exists(path): continue

        play_video(path)
        win.flip() 
        
        opts = ["A. 积极 (Positive)", "B. 中性 (Neutral)", "C. 消极 (Negative)"]
        choice = get_rating(f"视频 [{vid_file}] 评价", opts, {'a':0, 'b':1, 'c':2})
        log_data("视频评分", vid_file, choice, 0)
        
        run_rest(TIME_REST)

    c1 = get_rating("【总评】整体积极程度", ["A. 非常积极", "B. 一般", "C. 不够积极"], {'a':0, 'b':1, 'c':2})
    log_data("总评", "ALL", c1, 0)

    c2 = get_rating("【总评】广告类型偏好", ["F. 时尚广告", "S. 运动广告"], {'f':0, 's':1})
    pref_map = {'F': '时尚广告', 'S': '运动广告'}
    log_data("偏好", "ALL", pref_map.get(c2, c2), 0)

    visual.TextStim(win, text="实验结束\n数据已保存", font=FONT_MAIN, height=0.05, color=COLOR_ACCENT).draw()
    win.flip()
    core.wait(3)

except Exception as e:
    print(f"Error: {e}")
    try: cv2.destroyAllWindows()
    except: pass
finally:
    win.close()
    core.quit()