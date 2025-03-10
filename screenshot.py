from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import pyautogui
from PIL import Image
import io
import win32clipboard
from ppt_handler import add_slide_and_content  # 添加这行导入语句

def take_screenshots(df, ppt_path, title_text, update_progress, update_link_status, root):
    """
    执行网页截图并添加到PPT的主要函数
    """
    chrome_options = Options()
    chrome_options.add_argument('--start-maximized')  # 设置启动时最大化
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), 
                           options=chrome_options)
    driver.maximize_window()  # 确保窗口最大化
    
    time.sleep(3)
    total = len(df)
    
    # 创建临时目录用于保存截图
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(ppt_path)), "temp_screenshots")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    try:
        for i, row in df.iterrows():
            update_progress(i, total)
            
            # 创建新幻灯片并添加内容
            slide = add_slide_and_content(
                row['媒体名称'],
                row['发布时间'],
                title_text,
                row['链接']
            )
            
            if not slide:
                continue
            
            # 访问网页并截图
            driver.get(row['链接'])
            time.sleep(3)
            
            # 使用临时目录保存截图
            screenshot_path = os.path.join(temp_dir, f"temp_screenshot_{i}.png")
            driver.save_screenshot(screenshot_path)
            
            # 将截图添加到幻灯片
            if os.path.exists(screenshot_path):
                # 添加截图到幻灯片（缩小到三分之一）
                left = 72  # 1英寸
                top = 180  # 2.5英寸（文本框下方）
                width = 240  # 原始宽度的三分之一
                height = 180  # 保持宽高比
                slide.Shapes.AddPicture(
                    screenshot_path, 
                    False, True, 
                    left, top, 
                    width, height
                )
                os.remove(screenshot_path)
            
            update_link_status(i)
            root.update()
            
    finally:
        # 清理临时文件
        if os.path.exists(temp_dir):
            for file in os.listdir(temp_dir):
                try:
                    os.remove(os.path.join(temp_dir, file))
                except:
                    pass
            try:
                os.rmdir(temp_dir)
            except:
                pass
        driver.quit()

def copy_to_clipboard(image_path):
    """将图片复制到剪贴板"""
    image = Image.open(image_path)
    output = io.BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def paste_to_ppt(row, title_text):
    """将截图粘贴到PPT"""
    # 点击幻灯片中间位置
    pyautogui.click(x=500, y=500)
    time.sleep(0.5)
    
    # 粘贴图片
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    
