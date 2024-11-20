import os
import time

import DrissionPage
import requests
import pandas as pd
import re
import openpyxl
import tkinter as tk
import tkinter.messagebox
import cv2
import numpy as np
import pyautogui
from threading import Thread

""" bilibili 爬虫 """
# pyinstaller -F -w BilibiliCrawler.py --icon = logo。ico



def rule_dir(check_dir: str) -> str:
    """
    document: 本函数功能为将路径规范化, 不合法的字符被删除
    param check_dir: 需要进行规范化的路径
    return: 规范化后的路径
    example: rule_dir(r"E:\exam?ple")
    """
    import os
    import re
    intab = r'[?*/\|:><"]'
    path_tier_list = os.path.normpath(check_dir).split(os.sep)
    print(path_tier_list)
    for index in range(len(path_tier_list)):
        path_tier_list[index] = re.sub(intab, "",
                                       path_tier_list[index]).strip() if index > 0 else f"{path_tier_list[index]}\\"
        print(path_tier_list)
        while path_tier_list[index][-1] == ".":
            path_tier_list[index] = path_tier_list[index][0:-1]
    return os.path.join(*(base_name for base_name in path_tier_list))


def bilibili_crawler(key_words, output_path):
    edge_path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'  # 请改为你电脑内Chrome可执行文件路径
    co = DrissionPage.ChromiumOptions().set_browser_path(edge_path)  # 设置浏览器配置, 浏览器路径

    page = DrissionPage.ChromiumPage(co)  # 打开浏览器
    # page = DrissionPage.ChromiumPage()

    KEY_WORDS = key_words.split(",")  # ["RPA", "大数据"]
    main_tab = page.new_tab()
    new_data_list = []
    OUTPUT_PATH = output_path # r"D:\测试\bilibili数据.xlsx"


    for key_word in KEY_WORDS:  # 遍历每个关键词
        page_number = 0  # 查询的当前页数
        while True:  # 不断进行翻页查找, 直到查找不到数据之后会进行 break
            video_count = 0  # 设置初始默认视频数量为 0
            page_number += 1  # 页数自增
            # 获取 url
            if page_number == 1:
                url = f"http://search.bilibili.com/all?keyword={key_word}"
            else:
                url = f"http://search.bilibili.com/all?keyword={key_word}&page={page_number}"
            main_tab.get(url)  # 跳转到 url 链接

            videos_ele = main_tab.ele('xpath://*[@id="i_cecream"]//div[@class="video-list row"]', timeout=2)  # 当前页所有视频的父 div 元素
            try:
                video_count = len(videos_ele.texts())  # 当前页的视频数量
            except Exception as e:
                pass
            if video_count == 0:  # 如果获取不到视频数量, 代表已经没有数据了, 则 break 当前关键词
                break

            for video_index in range(video_count):  # 遍历循环视频数量次数  # video_count
                video_id = video_index + 1  # 视频的 id 为 视频的 index + 1
                # video_ele = main_tab.ele(f'xpath://*[@id="i_cecream"]/div/div[2]/div[2]/div/div/div/div[2]/div/div[{video_id}]/div/div[2]/div/div/a')  # 查找视频的 a 标签 元素
                video_ele = main_tab.ele(f'xpath://*[@class="video-list row"]/div[{video_id}]/div[@class="bili-video-card"]/div[@class="bili-video-card__wrap __scale-wrap"]/a')
                video_href = video_ele.link  # 获取视频 a 标签元素下的 href
                # video_name = video_ele.ele('xpath:h3[@class="bili-video-card__info--tit"]').title
                video_name = main_tab.ele(f'xpath://*[@class="video-list row"]/div[{video_id}]/div[@class="bili-video-card"]/div[@class="bili-video-card__wrap __scale-wrap"]/div[@class="bili-video-card__info __scale-disable"]/div[@class="bili-video-card__info--right"]/a/h3').title

                # video_worker = main_tab.ele(f'xpath://*[@id="i_cecream"]/div/div[2]/div[2]/div/div/div/div[2]/div/div[{video_id}]/div/div[2]/div/div/p/a/span[1]').text
                video_worker = main_tab.ele(f'xpath://*[@class="video-list row"]/div[{video_id}]/div[@class="bili-video-card"]/div[@class="bili-video-card__wrap __scale-wrap"]/div[@class="bili-video-card__info __scale-disable"]/div/p/a/span[1]').text

                # video_date = main_tab.ele(f'xpath://*[@id="i_cecream"]/div/div[2]/div[2]/div/div/div/div[2]/div/div[{video_id}]/div/div[2]/div/div/p/a/span[2]').text
                video_date = main_tab.ele(f'xpath://*[@class="video-list row"]/div[1]/div[@class="bili-video-card"]/div[@class="bili-video-card__wrap __scale-wrap"]/div[@class="bili-video-card__info __scale-disable"]/div/p/a/span[2]').text

                video_tab = page.new_tab()  # 生成一个用于打开新的视频的 tab 页
                video_tab.get(video_href)  # 跳转到 url 链接
                is_teach = False
                while True:
                    video_tab.scroll.to_bottom()  # 滚动到最底部
                    time.sleep(1)  # 等待加载
                    # 判断是否加载完所有评论
                    try:
                        try:
                            # 判断是否是买课的
                            ele = video_tab.ele('xpath://*[@class="comment-tab tab-item flex-align-center"]/p', timeout=2)
                            # print(ele)
                            # print(ele.text)
                            ele.click()
                            is_teach = True
                            break
                        except Exception as e:
                            pass
                        # 加载完所有评论时才会出现的 end 标签
                        end_ele = video_tab.ele('xpath://*[@id="comment"]/div[@class="comment"]/bili-comments').shadow_root.ele('xpath:div[@id="end"]/div[@class="bottombar"]', timeout=2)
                        if end_ele.text == "没有更多评论":
                            break  # 出现 end 标签, 跳出
                    except Exception as e:
                        continue  # 没有出现 end 标签, 继续
                if is_teach:
                    video_tab.close()  # 关闭视频的 tab 页
                    continue  # 如果是卖课的就不抓了, 是另一种布局, 懒得适配了

                # feed 元素, 所有的评论都在 feed 元素里面
                feed_ele = video_tab.ele('xpath://*[@id="comment"]/div[@class="comment"]/bili-comments').shadow_root.ele('xpath:div[@id="contents"]/div[@id="feed"]')
                comment_index = 0  # 评论的下标
                while True:
                    try:
                        data = {}
                        comment_index += 1  # 评论下标实际从 1 开始
                        comment_ele = feed_ele.ele(f'xpath:bili-comment-thread-renderer[{comment_index}]', timeout=2).shadow_root.ele('xpath:bili-comment-renderer[@id="comment"]').shadow_root.ele('xpath:div[@id="body"]/div[@id="main"]')
                        # 评论人昵称
                        name_ele = comment_ele.ele('xpath:div[@id="header"]/bili-comment-user-info').shadow_root.ele('xpath:div[@id="info"]/div[@id="user-name"]/a')
                        content_ele = comment_ele.ele('xpath:div[@id="content"]/bili-rich-text').shadow_root.ele('xpath:p[@id="contents"]')  # /span
                        time_ele = comment_ele.ele('xpath:div[@id="footer"]/bili-comment-action-buttons-renderer').shadow_root.ele('xpath:div[@id="pubdate"]')
                        thumbs_up_ele = comment_ele.ele('xpath:div[@id="footer"]/bili-comment-action-buttons-renderer').shadow_root.ele('xpath:div[@id="like"]/button/span[@id="count"]')
                        thumbs_up_count = thumbs_up_ele.text if thumbs_up_ele.text else 0

                        # print(video_name)  # 视频名称
                        # print(video_worker)  # 视频作者
                        # print(video_date)  # 视频日期
                        # print(name_ele.text)  # 评论人
                        # print(content_ele.text)  # 评论内容
                        # print(time_ele.text)  # 评论时间
                        # print(thumbs_up_count)  # 点赞数

                        data["视频名称"] = video_name
                        data["视频作者"] = video_worker
                        data["视频日期"] = video_date
                        data["评论人"] = name_ele.text
                        data["评论内容"] = content_ele.text
                        data["评论时间"] = time_ele.text
                        data["点赞数"] = thumbs_up_count
                        new_data_list.append(data)
                    except Exception as e:
                        # print(f"一共{comment_index-1}条评论")
                        break
                pd.DataFrame.from_records(new_data_list).to_excel(OUTPUT_PATH, index=False)
                # time.sleep(100)
                video_tab.close()  # 关闭视频的 tab 页


class GUI:
    def __init__(self):
        # self.is_over = False
        self.root = tk.Tk()
        # 设置窗口信息
        WIDTH = 400
        HEIGHT = 300
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (WIDTH / 2)
        y = (screen_height / 2) - (HEIGHT / 2)
        self.root.geometry(f"{WIDTH}x{HEIGHT}+{int(x)}+{int(y)}")  # 窗口尺寸, 坐标
        self.root.title("飞书自动考勤机器人")

    def set_grid(self):
        # 建立布局
        for i in range(2):  # 3列
            self.root.columnconfigure(i, weight=1)
        for i in range(6):  # 6行
            self.root.rowconfigure(i, weight=1)

        # 创建标签
        label1 = tk.Label(self.root, text="关键词(用英文逗号','隔开):")
        label1.grid(row=1, column=0)
        label2 = tk.Label(self.root, text="保存路径:")
        label2.grid(row=2, column=0)

        # 创建输入框
        self.entry1 = tk.Entry(self.root, textvariable=tk.StringVar(value="RPA,大数据"))
        self.entry1.grid(row=1, column=1)
        self.entry2 = tk.Entry(self.root, textvariable=tk.StringVar(value=r"D:\测试\bilibili数据.xlsx"))
        self.entry2.grid(row=2, column=1)

        # 创建按钮并绑定事件
        button1 = tk.Button(self.root, text="开始执行", command=self._func_start)
        button1.grid(row=4, column=0)
        button2 = tk.Button(self.root, text="退出程序", command=self.root.destroy)
        button2.grid(row=4, column=1)

    def display(self):
        self.set_grid()
        self.root.mainloop()
        # self.get_video()

    def _func_start(self):
        global START_DATE, END_DATE, EXCEL_DOWNLOAD_PATH, EXCEL_OUTPUT_PATH
        KEY_WORDS = self.entry1.get()
        OUTPUT_PATH = self.entry2.get()
        if not os.path.isdir(os.path.dirname(OUTPUT_PATH)):
            tk.messagebox.showinfo(title='提示', message=f'不存在目录:\n{os.path.dirname(OUTPUT_PATH)}')
            return False


        try:
            bilibili_crawler(key_words=KEY_WORDS, output_path=OUTPUT_PATH)
            tk.messagebox.showinfo(title='提示', message='执行成功')
        except Exception as e:
            tk.messagebox.showinfo(title='提示', message=f'执行失败\n{e}')


        # pattern = r'^\d{4}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$'  # 正则表达式，匹配一个字母
        # if bool(re.match(pattern, START_DATE)) and bool(re.match(pattern, END_DATE)):
        #     EXCEL_DOWNLOAD_PATH = os.path.join(EXCEL_DOWNLOAD_DIR, f"{START_DATE}-{END_DATE}考勤.xlsx")
        #     EXCEL_OUTPUT_PATH = os.path.join(EXCEL_DOWNLOAD_DIR, f"{START_DATE}-{END_DATE}统计.xlsx")
        #     try:
        #         download_excel()
        #         change_excel()
        #         tk.messagebox.showinfo(title='提示', message='执行成功')
        #
        #     except Exception as e:
        #         tk.messagebox.showinfo(title='提示', message=f'执行失败\n{e}')
        #     global IS_OVER
        #     IS_OVER = True
        # else:
        #     tk.messagebox.showinfo(title='提示', message='日期格式不正确, 请重新输入\n\n日期格式: 20220101')
# print(os.path.dirname(r"D:\测试\33"))
# print()
gui_obj = GUI()
gui_obj.display()
