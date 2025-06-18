# -*- coding:utf-8 -*-
import json
import os
import re
import time
import io
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pyzipper


class XiaohongshuDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("仅供学习")
        self.root.geometry("800x600")

        # 下载配置
        self.config = {
            'HEADLESS': tk.BooleanVar(value=True),
            'CHROME_DRIVER_PATH': tk.StringVar(),
            'WAIT_TIME': tk.IntVar(value=10),
            'ZIP_PATH': tk.StringVar(value="小红书照片.zip"),
            'ZIP_PASSWORD': tk.StringVar(value="zippassword3332")
        }

        # 下载状态
        self.excel_file = tk.StringVar()
        self.is_running = False
        self.progress_value = tk.IntVar()
        self.progress_text = tk.StringVar(value="准备就绪")
        self.driver = None

        # 创建UI
        self.create_widgets()

    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择部分
        file_frame = ttk.LabelFrame(main_frame, text="Excel文件选择", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        ttk.Label(file_frame, text="Excel文件路径:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.excel_file, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="浏览...", command=self.browse_file).grid(row=0, column=2)

        # 设置部分
        settings_frame = ttk.LabelFrame(main_frame, text="下载设置", padding="10")
        settings_frame.pack(fill=tk.X, pady=5)

        ttk.Checkbutton(settings_frame, text="无头模式", variable=self.config['HEADLESS']).grid(row=0, column=0,
                                                                                                sticky=tk.W)

        ttk.Label(settings_frame, text="等待时间(秒):").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(settings_frame, textvariable=self.config['WAIT_TIME'], width=10).grid(row=2, column=1, sticky=tk.W)

        # 进度条部分
        progress_frame = ttk.LabelFrame(main_frame, text="下载进度", padding="10")
        progress_frame.pack(fill=tk.X, pady=5)

        ttk.Label(progress_frame, textvariable=self.progress_text).pack(anchor=tk.W)
        ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=500,
                        mode='determinate', variable=self.progress_value).pack(fill=tk.X, pady=5)

        # 日志部分
        log_frame = ttk.LabelFrame(main_frame, text="下载日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 按钮部分
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="开始下载", command=self.start_download).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="停止", command=self.stop_download).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.quit_app).pack(side=tk.RIGHT)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.excel_file.set(file_path)

    def log_message(self, message):
        """将消息添加到日志文本框"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # 自动滚动到底部
        self.root.update_idletasks()  # 更新UI

    def start_download(self):
        if not self.excel_file.get():
            messagebox.showerror("错误", "请先选择Excel文件")
            return

        if not os.path.exists(self.excel_file.get()):
            messagebox.showerror("错误", "Excel文件不存在")
            return

        try:
            # 测试是否能读取Excel文件
            df = pd.read_excel(self.excel_file.get(), sheet_name='data', engine='openpyxl')
            if df.empty:
                messagebox.showerror("错误", "Excel文件为空")
                return
        except Exception as e:
            messagebox.showerror("错误", f"无法读取Excel文件:\n{str(e)}")
            return

        if self.is_running:
            messagebox.showwarning("警告", "下载任务已在运行")
            return

        # 清空日志
        self.log_text.delete(1.0, tk.END)

        self.is_running = True
        self.progress_value.set(0)
        self.progress_text.set("正在准备下载...")
        self.log_message("=== 开始下载任务 ===")

        # 在新线程中运行下载任务
        download_thread = threading.Thread(
            target=self.run_download_task,
            daemon=True
        )
        download_thread.start()

    def stop_download(self):
        if self.is_running:
            self.is_running = False
            self.progress_text.set("正在停止下载...")
            self.log_message("用户请求停止下载...")
        else:
            messagebox.showinfo("信息", "没有正在运行的下载任务")

    def quit_app(self):
        if self.is_running:
            if messagebox.askokcancel("退出", "下载任务正在进行中，确定要退出吗？"):
                self.is_running = False
                if self.driver:
                    self.driver.quit()
                self.root.quit()
        else:
            if self.driver:
                self.driver.quit()
            self.root.quit()

    def update_progress(self, value, text):
        self.progress_value.set(value)
        self.progress_text.set(text)
        self.root.update_idletasks()

    def sanitize_filename(self, name):
        """将非法文件名字符替换为下划线，以创建合法的文件夹名。"""
        return re.sub(r'[\/\\\:\*\?\"\<\>\|]', '_', name).strip()

    def download_image_to_zip(self, zipf, url, filepath_in_zip):
        """下载图片并直接写入到压缩包中，返回是否成功。"""
        try:
            self.log_message(f"正在下载: {url}")
            res = requests.get(url, stream=True, timeout=10)
            if res.status_code == 200:
                # 使用内存缓冲区存储图片数据
                img_data = io.BytesIO()
                for chunk in res.iter_content(chunk_size=8192):
                    if chunk:
                        img_data.write(chunk)
                        if not self.is_running:
                            return False

                # 将图片数据写入压缩包
                img_data.seek(0)
                zipf.writestr(filepath_in_zip, img_data.read())
                self.log_message(f"下载成功: {url}")
                return True
        except Exception as e:
            self.log_message(f'下载异常: {url}, 错误: {e}')
        return False

    def get_images_from_xhs(self, url, driver, wait_time=10):
        """
        使用Selenium访问小红书链接，提取详情页的图片链接列表。
        利用页面中包含的window.__INITIAL_SSR_STATE__脚本获取imageList。
        """
        images = []
        try:
            self.log_message(f"正在访问小红书页面: {url}")
            driver.get(url)
            # 等待页面加载并包含INITIAL_SSR_STATE脚本
            try:
                WebDriverWait(driver, wait_time).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//script[contains(text(), 'window.__INITIAL_SSR_STATE__')]"))
                )
            except Exception:
                # 如果未找到，也等待短暂时间后尝试获取
                time.sleep(3)
            # 查找包含INITIAL_SSR_STATE脚本的元素
            script_elem = driver.find_element(By.XPATH, "//script[contains(text(), 'window.__INITIAL_SSR_STATE__')]")
            script_content = script_elem.get_attribute('innerHTML')
            # 提取JSON字符串部分（去除前缀和末尾分号）
            try:
                json_str = script_content.split('=', 1)[1].strip().rstrip(';')
            except Exception:
                self.log_message("无法解析INITIAL_SSR_STATE内容")
                return images
            # 替换undefined为null，以便解析
            json_str = json_str.replace('undefined', 'null')
            data = json.loads(json_str)
            # 提取imageList
            note_info = data.get('NoteView', {}).get('noteInfo')
            if note_info:
                for img in note_info.get('imageList', []):
                    url_part = img.get('url', '').strip()
                    if url_part:
                        # 处理URL编码
                        url_clean = url_part.replace('\\u002F', '/')
                        # 添加https前缀
                        if not url_clean.startswith('http'):
                            url_clean = 'https:' + url_clean
                        images.append(url_clean)
                self.log_message(f"从小红书页面获取到 {len(images)} 张图片")
        except Exception as e:
            self.log_message(f'使用Selenium获取图片链接失败: {e}')
        return images

    def run_download_task(self):
        try:
            # 读取Excel文件
            df = pd.read_excel(self.excel_file.get(), sheet_name='data', engine='openpyxl')
            total_items = len(df)
            self.log_message(f"Excel文件中找到 {total_items} 条记录")

            # 初始化Selenium WebDriver
            options = Options()
            if self.config['HEADLESS'].get():
                options.add_argument('--headless')
                self.log_message("使用无头模式启动浏览器")

            self.driver = webdriver.Chrome(options=options)
            self.log_message("浏览器已启动")

            # 创建加密压缩包
            zip_path = self.config['ZIP_PATH'].get()
            with pyzipper.AESZipFile(zip_path,
                                     'w',
                                     compression=pyzipper.ZIP_DEFLATED,
                                     encryption=pyzipper.WZ_AES) as zipf:
                zipf.setpassword(self.config['ZIP_PASSWORD'].get().encode('utf-8'))
                self.log_message(f"创建加密压缩包: {zip_path}")

                # 遍历每一行数据
                for idx, row in df.iterrows():
                    if not self.is_running:
                        break

                    # 更新进度
                    progress = int((idx + 1) / total_items * 100)
                    title = str(row.get('标题', '')).strip()
                    self.update_progress(progress, f"正在处理第 {idx + 1}/{total_items} 条: {title}")

                    content_imgs = str(row.get('内容图片', '')).strip()  # 内容图片（可能多张）

                    # 生成文件夹名
                    folder_name = self.sanitize_filename(title) or f"item_{idx}"
                    self.log_message(f"\n正在处理: {title} (保存到文件夹: {folder_name})")

                    # 下载内容图片（可能多张，用逗号/分号分隔）
                    if content_imgs and self.is_running:
                        img_urls = re.split(r'\s+', content_imgs.strip())
                        self.log_message(f"发现 {len(img_urls)} 张内容图片")
                        for i, img_url in enumerate(img_urls):
                            if not self.is_running:
                                break

                            img_url = img_url.strip()
                            if not img_url:
                                continue

                            ext = os.path.splitext(img_url)[1] or '.jpg'
                            save_path = os.path.join(folder_name, f"内容图_{i + 1}{ext}")
                            if self.download_image_to_zip(zipf, img_url, save_path):
                                self.log_message(f'内容图片 {i + 1} 下载成功: {img_url}')
                            else:
                                self.log_message(f'内容图片 {i + 1} 下载失败: {img_url}')

            if self.is_running:
                self.update_progress(100, "下载完成!")
                self.log_message("\n=== 下载任务完成 ===")
                messagebox.showinfo("完成", f"所有图片已下载完成并保存到 {self.config['ZIP_PATH'].get()}")
            else:
                self.update_progress(0, "下载已停止")
                self.log_message("\n=== 下载任务已停止 ===")

            self.is_running = False

        except Exception as e:
            self.update_progress(0, f"错误: {str(e)}")
            self.log_message(f"\n!!! 发生错误: {str(e)}")
            messagebox.showerror("错误", f"下载过程中发生错误:\n{str(e)}")
            self.is_running = False
        finally:
            if self.driver:
                self.driver.quit()
                self.driver = None
                self.log_message("浏览器已关闭")


if __name__ == "__main__":
    root = tk.Tk()
    app = XiaohongshuDownloaderApp(root)
    root.mainloop()