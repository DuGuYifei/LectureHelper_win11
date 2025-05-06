import os

import psutil
import subprocess
import time
from random import randint
from datetime import datetime, timedelta
from PIL import ImageGrab

import requests
import uiautomation as auto
from pynput import mouse
import imagehash


def on_click(x, y, button, pressed):
    if pressed:
        print(f"鼠标点击位置：({x}, {y})")
        # 停止监听
        return False


LIVECAP_EXE = r"C:\Windows\System32\LiveCaptions.exe"
POLL_INTERVAL = timedelta(seconds=120)  # seconds
WRITE_TO_FILE_INTERVAL = timedelta(seconds=5)  # seconds


def screenshot():
    img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
    return img


def random_time(start_time, end_time):
    return timedelta(seconds=randint(start_time, end_time))


def kill_all_livecaptions():
    """Terminate any running LiveCaptions.exe processes."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and proc.info['name'].lower() == 'livecaptions.exe':
            try:
                proc.kill()
            except psutil.NoSuchProcess:
                pass


def get_livecaptions_proc():
    """Get the LiveCaptions.exe process."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and proc.info['name'].lower() == 'livecaptions.exe':
            return proc
    return None


def start_livecaptions():
    """Start a new LiveCaptions.exe (hidden window)."""
    return subprocess.Popen(
        [LIVECAP_EXE],
        creationflags=subprocess.CREATE_NO_WINDOW
    )


def google_translate_web(text, target_lang='zh-CN'):
    """
    使用 Google Translate 非官方接口翻译文本。

    参数:
        text (str): 要翻译的文本
        target_lang (str): 目标语言（例如 'zh-CN', 'en', 'de' 等）

    返回:
        str: 翻译后的文本
    """
    base_url = "https://clients5.google.com/translate_a/t"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Referer": "https://translate.google.com/"
    }

    params = {
        "client": "dict-chrome-ex",
        "sl": "auto",  # 源语言自动检测
        "tl": target_lang,
        "q": text
    }

    response = requests.get(base_url, headers=headers, params=params)

    if response.status_code == 200:
        try:
            data = response.json()
            # 返回结构如：["你好", "hello", ..., ...]
            return data[0]
        except Exception as e:
            return f"解析失败: {e}"
    else:
        return f"请求失败，状态码: {response.status_code}"


def is_slide_changed(current_img, last_hash, threshold=5):
    processed = current_img.convert('L').resize((256, 256))
    current_hash = imagehash.phash(processed)

    if last_hash is None:
        last_hash = current_hash
        return True, last_hash

    changed = abs(current_hash - last_hash) > threshold
    if changed:
        last_hash = current_hash  # 更新哈希值
    return changed, last_hash


def main(is_screenshot=False):
    # 1) Kill old instances
    kill_all_livecaptions()
    time.sleep(0.2)

    # 2) Launch Live Captions
    proc = start_livecaptions()

    # 3) Give it a moment to spin up and register its UI
    time.sleep(1.0)

    # 测试用-获取当前的live captions进程
    # proc = get_livecaptions_proc()
    # if proc is None:
    #     print("❌ Cannot find Live Captions process.")
    #     return

    # 4) Attach to the LiveCaptions window
    window = auto.WindowControl(
        searchDepth=1,
        ProcessId=proc.pid,
        ClassName='LiveCaptionsDesktopWindow'
    )
    if not window.Exists(5):
        print("❌ Cannot find Live Captions window.")
        return

    # 5) Find the scroll-viewer pane that holds the text
    scroll = window.Control(
        AutomationId='CaptionsScrollViewer'
    )
    while not scroll.Exists(5):
        print("❌ Cannot find CaptionsScrollViewer element. May because there is no audio playing. \n- 可能是因为没有音频播放。将循环等待...")
        time.sleep(1.0)

    print("✅ Attached! Waiting for captions…")
    datetime_now = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    # 创建文件夹
    if not os.path.exists(datetime_now):
        os.makedirs(datetime_now)

    try:
        print(scroll.Name)
        timer_write_file = datetime.now()
        timer_poll = datetime.now()
        timer_translate = datetime.now()
        last_text_poll = ""
        last_index_poll = 0
        last_text_last_part = ""
        last_text_write = ""
        previous_hash = None
        index_count_when_smaller_than_maximum_caption = 0
        while True:
            # 进行请求
            # if timer_poll + POLL_INTERVAL <= datetime.now():
            #     text = scroll.Name  # this is the current caption line(s)
            #     if text != last_text_poll:
            #         print(f"{text[last_index_poll:]}\n---")
            #         last_index_poll = len(text) - 10 if len(text) > 10 else 0
            #         last_text_poll = text
            #         pass
            #     timer_poll = datetime.now()

            # 将text写入文件，续写
            if timer_write_file + WRITE_TO_FILE_INTERVAL <= datetime.now():
                text = scroll.Name  # this is the current caption line(s)
                if text == last_text_write:
                    print("没有新的内容，跳过写入文件")
                    continue
                last_text_write = text
                with open(f"{datetime_now}/live_caption_{datetime_now}.md", "a", encoding="utf-8") as f:
                    # 只写入新增的内容
                    # 找到text的last_text_last_part，然后续写
                    # print(f"写入文件：{text[text.index(last_text_last_part) + len(last_text_last_part) - 10:]}")
                    temp_index = len(text)
                    if temp_index > 400:
                        print("大于400，写入文件")
                        f.write(text[text.index(last_text_last_part) + len(last_text_last_part):])
                    else:
                        print("小于400，写入文件")
                        f.write(text[index_count_when_smaller_than_maximum_caption:])
                        index_count_when_smaller_than_maximum_caption = temp_index
                    # 添加图片md语法
                    if is_screenshot:
                        # 截图
                        screenshot_img = screenshot()
                        # 判断是否有变化
                        is_img_changed, previous_hash = is_slide_changed(screenshot_img, previous_hash)
                        if is_img_changed:
                            print("截图有变化，保存截图")
                            current_image_filename = f"live_caption_{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.png"
                            screenshot_img.save(f"{datetime_now}/{current_image_filename}")
                            # 添加图片md语法
                            f.write(f"\n\n![{current_image_filename}]({current_image_filename})\n")
                    last_text_last_part = text[len(text) - 100:len(text) - 30] if len(text) > 200 else ""
                    timer_write_file = datetime.now()

            # 进行翻译
            if timer_translate + random_time(translate_delta_int_low, translate_delta_int_high) <= datetime.now():
                text = scroll.Name  # this is the current caption line(s)
                start_translation_index = len(text) - 500 if len(text) > 500 else 0
                translated_text = google_translate_web(text[start_translation_index:], target_lang="zh-CN")
                print(f"翻译结果{translated_text}\n---")
                timer_translate = datetime.now()
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nInterrupted, exiting.")
    finally:
        # clean up
        try:
            proc.kill()
        except Exception:
            pass


if __name__ == '__main__':
    print("输入后按回车键")
    screenshot_input = input("是否截图？(y/n)：")
    write_to_file_delta_int = input("将字幕和截图存进文件的间隔时间设置（电脑越好可以越快）")
    WRITE_TO_FILE_INTERVAL = timedelta(seconds=int(write_to_file_delta_int))
    print("翻译间隔时间设置，（建议别太快，google翻译，封IP）")
    translate_delta_int_low = input("随机翻译间隔下限（秒）建议6以上：")
    translate_delta_int_low = int(translate_delta_int_low)
    translate_delta_int_high = input("随机翻译间隔上限（秒）：")
    translate_delta_int_high = int(translate_delta_int_high)
    x1 = y1 = x2 = y2 = 0
    if screenshot_input.lower() == 'y':
        print("请点击截图区域的左上角")
        # 等待鼠标点击
        # 开始监听鼠标点击
        with mouse.Listener(on_click=on_click) as listener:
            listener.join()
        # 获取鼠标点击位置
        x1, y1 = auto.GetCursorPos()
        print(f"左上角坐标：({x1}, {y1})")
        print("Please click the right bottom corner of screenshot area - 请点击截图区域的右下角")
        # 等待鼠标点击
        # 开始监听鼠标点击
        with mouse.Listener(on_click=on_click) as listener:
            listener.join()
        # 获取鼠标点击位置
        x2, y2 = auto.GetCursorPos()

    main(is_screenshot=(screenshot_input.lower() == 'y'))
