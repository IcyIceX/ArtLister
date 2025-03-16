import os
import sys
import json
import yaml
import FreeSimpleGUI as sg
from markitdown import MarkItDown
import pandas as pd
from openai import OpenAI
from pathlib import Path
import time
import traceback

class ArtLister:
    def __init__(self):
        # 获取应用程序路径
        self.application_path = self.get_application_path()
        
        # 初始化其他属性
        self.config = self.load_config()
        self.output_dir = self.get_output_dir()
        self.md = MarkItDown()
        self.window = None
        self.api_key = ""
        self.log_messages = []

    def get_application_path(self):
        """获取应用程序路径"""
        if getattr(sys, 'frozen', False):
            # 如果是打包后的可执行文件
            return os.path.dirname(sys.executable)
        else:
            # 如果是脚本运行
            return os.path.dirname(os.path.abspath(__file__))
        
    def load_config(self):
        """加载配置文件"""
        config_path = os.path.join(self.application_path, 'config.yaml')
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            self.log(f"Error loading config: {e}")
            # 返回默认配置
            return {
                'output': {
                    'directory': 'output',
                    'excel_filename': 'art_list.xlsx'
                }
            }        
        
    def get_output_dir(self):
        """获取输出目录路径"""
        output_dir = Path(self.application_path) / self.config['output']['directory']
        output_dir.mkdir(exist_ok=True)
        return output_dir  

    def get_unique_filename(self, base_path):
        """生成唯一的文件名，避免覆盖"""
        path = Path(base_path)
        counter = 1
        while path.exists():
            path = Path(f"{path.parent}/{path.stem}({counter}){path.suffix}")
            counter += 1
        return str(path)          

    def log(self, message):
        """记录日志"""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        log_entry = f"[{timestamp}] {message}"
        self.log_messages.append(log_entry)
        
        # 如果窗口已创建，更新日志显示
        if self.window and '-LOG-' in self.window.AllKeysDict:
            self.window['-LOG-'].update('\n'.join(self.log_messages))
            # 自动滚动到最新日志
            self.window['-LOG-'].set_vscroll_position(1.0)

    def create_gui(self):
        """创建GUI界面"""
        # 设置主题（如果支持）
        if hasattr(sg, 'theme'):
            sg.theme('LightGrey1')
        else:
            print("警告：PySimpleGUI版本可能不支持theme方法，使用默认主题。")

        # 定义布局
        layout = [
            [sg.Text('欢迎使用美术清单助手', font=('Any', 16))],
            [sg.Text('请输入Deepseek密钥:', size=(15, 1)), 
             sg.Input(key='-API_KEY-', password_char='*', size=(45, 1))],
            [sg.Text('选择文件:', size=(15, 1)), 
             sg.Input(key='-FILE-', size=(35, 1)), 
             sg.FileBrowse(file_types=(("所有支持的文件", "*.xlsx *.doc *.docx *.txt"),))],
            [sg.Text('处理状态:', size=(15, 1)), 
             sg.Text('等待开始...', key='-STATUS-', size=(45, 1))],
            [sg.ProgressBar(100, orientation='h', size=(40, 20), key='-PROGRESS-')],
            [sg.Text('详细信息:')],
            [sg.Multiline(size=(60, 10), key='-LOG-', autoscroll=True, disabled=True)],
            [sg.Button('开始处理', key='-PROCESS-'), sg.Button('退出')]
        ]

        # 创建窗口
        self.window = sg.Window('ArtListerv1.0-dev-by-chopin', layout, finalize=True)
        
        # 初始日志
        self.log("应用程序已启动")
        self.log(f"工作目录: {self.application_path}")

    def clean_json_response(self, response_text):
        """清理API响应中的Markdown代码块标记"""
        self.log("开始清理JSON响应内容")
        
        # 移除开头的 ```json 和结尾的 ``` 标记
        cleaned_text = response_text.strip()
        if cleaned_text.startswith('```json'):
            cleaned_text = cleaned_text[7:]  # 移除 ```json
        elif cleaned_text.startswith('```'):
            cleaned_text = cleaned_text[3:]  # 移除 ```
            
        if cleaned_text.endswith('```'):
            cleaned_text = cleaned_text[:-3]  # 移除结尾的 ```
            
        cleaned_text = cleaned_text.strip()
        self.log(f"清理后的JSON长度: {len(cleaned_text)} 字符")
        return cleaned_text

    def process_file(self, file_path, apiKey):
        """处理文件的主要逻辑"""
        try:
            # 更新状态：开始转换文件
            self.update_status('正在转换文件格式...', 20)
            self.log(f"开始处理文件: {file_path}")
            
            # 使用markitdown转换文件
            self.log("使用MarkItDown转换文件格式")
            result = self.md.convert(file_path)
            markdown_content = result.text_content
            self.log(f"文件转换成功，文本长度: {len(markdown_content)} 字符")
            
            # 更新状态：开始调用AI
            self.update_status('正在调用LLM分析文本...', 40)
            
            # 初始化OpenAI客户端
            self.log("初始化DeepSeek API客户端")
            client = OpenAI(api_key=apiKey, base_url="https://api.deepseek.com/v1")
            
            # 读取prompt模板
            prompt_path = os.path.join(self.application_path, 'prompt.md')
            self.log(f"加载Prompt模板: {prompt_path}")
            try:
                with open(prompt_path, 'r', encoding='utf-8') as f:
                    prompt_template = f.read()
                self.log(f"Prompt模板加载成功，长度: {len(prompt_template)} 字符")
            except Exception as e:
                self.log(f"Prompt模板加载失败: {str(e)}")
                raise
            
            # 调用Deepseek API
            self.log("开始调用DeepSeek API，请耐心等待...")
            self.window.refresh()  # 刷新UI以显示最新日志
            
            try:
                # 添加超时和更多设置
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": prompt_template},
                        {"role": "user", "content": markdown_content}
                    ],
                    temperature=1.0,          # 控制创造性
                    max_tokens=4096,          # 限制响应长度
                    stream=False,             # 不使用流式处理
                    timeout=180               # 超时设置3分钟
                )
                self.log("DeepSeek API调用成功")
            except Exception as e:
                self.log(f"API调用失败: {str(e)}")
                raise
            
            # 更新状态：生成JSON
            self.update_status('正在解析API响应...', 60)
            
            # 解析API响应
            try:
                json_content = response.choices[0].message.content.strip()
                self.log(f"API响应长度: {len(json_content)} 字符")
                self.log("API响应内容前500字符: " + json_content[:500] + "...")

                # 清理JSON响应
                cleaned_json = self.clean_json_response(json_content)
                self.log("尝试解析清理后的JSON响应")
                
                # 检查JSON是否有效
                json_data = json.loads(cleaned_json)
                self.log("JSON解析成功")

                # 验证JSON结构
                if '场景道具清单' not in json_data:
                    self.log("警告: 未找到'场景道具清单'键")
                    self.log(f"可用的键: {list(json_data.keys())}")
                else:
                    self.log(f"找到场景道具清单，包含 {len(json_data['场景道具清单'])} 个项目")                

            except json.JSONDecodeError as e:
                self.log(f"JSON解析错误: {str(e)}")
                self.log("原始响应内容: " + json_content)
                self.log("清理后的内容: " + cleaned_json)
                raise
            
            # 保存JSON文件
            self.update_status('正在保存JSON文件...', 70)
            self.save_json(cleaned_json)
            
            # 更新状态：生成Excel
            self.update_status('正在生成Excel...', 80)
            self.generate_excel(cleaned_json)  # 使用清理后的JSON
            
            # 完成处理
            self.update_status('处理完成！', 100)
            self.log("处理完成!")
            
            # 显示完成消息
            excel_path = self.get_excel_path()
            sg.popup(f'文件处理完成！\n输出文件位置：\n{excel_path}')
            
        except Exception as e:
            error_detail = traceback.format_exc()
            self.log(f"错误详情: {error_detail}")
            sg.popup_error(f'处理过程中出现错误：\n{str(e)}')
            self.update_status('处理失败', 0)

    def update_status(self, message, progress):
        """更新GUI状态和进度条"""
        self.window['-STATUS-'].update(message)
        self.window['-PROGRESS-'].update(progress)
        self.log(message)
        self.window.refresh()  # 立即刷新UI

    def save_json(self, json_content):
        """保存JSON文件"""
        try:
            json_path = self.output_dir / 'art_list.json'
            json_path = self.get_unique_filename(json_path)
            
            self.log(f"保存JSON文件到: {json_path}")
            
            # 如果输入是字符串，先解析为Python对象
            if isinstance(json_content, str):
                json_data = json.loads(json_content)
            else:
                json_data = json_content
                
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            self.log("JSON文件保存成功")
        except Exception as e:
            self.log(f"保存JSON文件失败: {str(e)}")
            raise

    def generate_excel(self, json_content):
        """生成Excel文件"""
        try:
            # 如果输入是字符串，先解析为Python对象
            if isinstance(json_content, str):
                json_data = json.loads(json_content)
            else:
                json_data = json_content
                
            self.log("开始生成Excel文件")
            
            if '场景道具清单' in json_data:
                df = pd.DataFrame(json_data['场景道具清单'])
                self.log(f"成功创建DataFrame，包含 {len(df)} 行数据")
            else:
                self.log("警告: 未找到'场景道具清单'键")
                # 尝试使用第一个可用的键
                first_key = list(json_data.keys())[0]
                self.log(f"尝试使用第一个键: {first_key}")
                df = pd.DataFrame(json_data[first_key])
                
            excel_path = self.output_dir / self.config['output']['excel_filename']
            excel_path = self.get_unique_filename(excel_path)
            self.log(f"生成Excel文件: {excel_path}")
            df.to_excel(excel_path, index=False)
            self.log("Excel文件生成成功")
        except Exception as e:
            self.log(f"生成Excel文件失败: {str(e)}")
            self.log(f"JSON数据结构: {json.dumps(json_data, indent=2)[:500]}...")  # 记录JSON结构
            raise

    def get_excel_path(self):
        """获取Excel文件路径"""
        return self.output_dir / self.config['output']['excel_filename']

    def run(self):
        """运行应用程序"""
        self.create_gui()
        
        while True:
            event, values = self.window.read(timeout=100)  # 添加超时参数使UI更响应
            
            if event in (sg.WIN_CLOSED, '退出'):
                break
                
            if event == '-PROCESS-':
                file_path = values['-FILE-']
                api_key = values['-API_KEY-']
                
                if not file_path:
                    sg.popup_error('请选择要处理的文件！')
                    self.log("错误: 未选择文件")
                    continue
                    
                if not api_key:
                    sg.popup_error('请输入API密钥！')
                    self.log("错误: 未输入API密钥")
                    continue
                
                # 禁用处理按钮，防止重复点击
                self.window['-PROCESS-'].update(disabled=True)
                self.process_file(file_path, api_key)
                # 重新启用处理按钮
                self.window['-PROCESS-'].update(disabled=False)
        
        self.log("应用程序关闭")
        self.window.close()

if __name__ == '__main__':
    app = ArtLister()
    app.run()
