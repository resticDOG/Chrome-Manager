<div align="center">

# Chrome 多窗口管理器

[![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB.svg?style=flat&logo=python&logoColor=white)](https://www.python.org)
[![Windows](https://img.shields.io/badge/Windows-10%2B-0078D6.svg?style=flat&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Chrome](https://img.shields.io/badge/Chrome-Latest-4285F4.svg?style=flat&logo=google-chrome&logoColor=white)](https://www.google.com/chrome/)
[![License](https://img.shields.io/badge/License-GPL%20v3-blue.svg)](LICENSE)



  <strong>作者：Devilflasher</strong>：<span title="No Biggie Community Founder"></span>
  [![X](https://img.shields.io/badge/X-1DA1F2.svg?style=flat&logo=x&logoColor=white)](https://x.com/DevilflasherX)
[![微信](https://img.shields.io/badge/微信-7BB32A.svg?style=flat&logo=wechat&logoColor=white)](https://x.com/DevilflasherX/status/1781563666485448736 "Devilflasherx")
 [![Telegram](https://img.shields.io/badge/Telegram-0A74DA.svg?style=flat&logo=telegram&logoColor=white)](https://t.me/devilflasher0) （欢迎加入微信群交流）
 

</div>

> [!IMPORTANT]
> ## ⚠️ 免责声明
> 
> 1. **本软件为开源项目，仅供学习交流使用，不得用于任何闭源商业用途**
> 2. **使用者应遵守当地法律法规，禁止用于任何非法用途**
> 3. **开发者不对因使用本软件导致的直接/间接损失承担任何责任**
> 4. **使用本软件即表示您已阅读并同意本免责声明**

## 工具介绍
Chrome 多窗口管理器是一款专门为 `NoBiggie社区` 准备的Chrome浏览器多窗口管理工具。它可以帮助用户轻松管理多个 Chrome 窗口，实现窗口批量打开、排列以及之间的同步操作，大大提高交互效率。

## 功能特性

- `批量管理功能`：一键打开/关闭单个、多个Chrome实例
- `智能布局系统`：支持自动网格排列和自定义坐标布局
- `多窗口同步控制`：实时同步鼠标/键盘操作到所有选定窗口
- `批量打开网页`：支持批量相同网页打开
- `快捷方式图标替换`：支持一键替换多个快捷方式图标（带序号的图标已准备在icon文件夹）
- `插件窗口同步`：支持弹出的插件窗口内的键盘和鼠标同步

## 环境要求

- Windows 10/11 (64-bit)
- Python 3.9+
- Chrome浏览器 最新

## 运行教程
### 方法一：打包成独立exe可执行文件（推荐）

如果你想自己打包程序，请按以下步骤操作：

1. **安装 Python 和依赖**
   ```bash
   # 安装 Python 3.9 或更高版本
   # 从 https://www.python.org/downloads/ 下载
   ```

2. **准备文件**
   - 确保目录里有以下文件：
     - chrome_manager.py（主程序）
     - build.py（打包脚本）
     - app.manifest（管理员权限配置）
     - app.ico（程序图标）

3. **运行打包脚本**
   ```bash
   # 在程序目录下运行：
   python build.py
   ```

4. **查找生成文件**
   - 打包完成后，在 `dist` 目录下找到 `chrome_manager.exe`
   - 双击运行 `chrome_manager.exe` 即可打开程序

### 方法二：从源码运行

1. **安装 Python**
   ```bash
   # 下载并安装 Python 3.9 或更高版本
   # 从 https://www.python.org/downloads/ 下载
   ```

2. **安装依赖包**
   ```bash
   # 打开命令提示符（CMD）并运行：
   pip install tkinter pywin32 keyboard mouse sv-ttk typing-extensions
   ```

3. **运行程序**
   ```bash
   # 在程序目录下运行：
   python chrome_manager.py
   ```

## 使用说明

### 前期准备


- 在您存放 Chrome 多开快捷方式的文件夹下，快捷方式的文件名应按照 `1.link`、`2.link`、`3.link`... 的格式命名。
- 同一个文件夹下建立 `Data` 文件夹，`Data` 文件夹下存放每个浏览器独立的数据文件，文件夹名应按照 `1`、`2`、`3`... 的格式命名。

```目录结构示例：
                                 多开chrome的目录

                                ├── 1.link
                                ├── 2.link
                                ├── 3.link
                                └── Data
                                    ├── 1
                                    ├── 2
                                    └── 3
```
- 浏览器快捷方式的目标参数因为：（请根据您的浏览器安装路径修改）
```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --user-data-dir="D:\chrom duo\Data\编号"
```

### 基本操作

1. **打开窗口**
   - 软件下方的"打开窗口"标签下，填入存放浏览器快捷方式的目录
   - "窗口编号"里填入想要打开的浏览器编号
   - 点击"打开窗口"按钮即可打开对应编号的chrome窗口

2. **导入窗口**
   - 点击"导入窗口"按钮导入当前打开的 Chrome 窗口
   - 在列表中选择要操作的窗口

3. **窗口排列**
   - 使用"自动排列"快速整理窗口
   - 或使用"自定义排列"设置详细的排列参数

4. **开启同步**
   - 选择一个主控窗口（点击"主控"列）
   - 选择要同步的从属窗口
   - 点击"开始同步"或使用设定的快捷键



## 注意事项

- 同步功能需要管理员权限
- 虽然理论上不会被杀毒软件误报或干扰，但请在报错时检查杀毒软件是否拦截相关功能
- 批量操作时注意系统资源占用

## 常见问题

1. **无法开启同步**
   - 检查是否以管理员身份运行
   - 确保已选择主控窗口

2. **窗口未正确导入**
   - 尝试重新点击"导入窗口"按钮
   - 
3. **滚动条同步幅度不同**
   - 目前的解决办法就是通过pageup和pagedown以及键盘上下左右键来调整同步幅度
   
  

## 更新日志

### v1.0
- 首次发布
- 实现基本的窗口管理和同步功能


## 许可证

本项目采用 GPL-3.0 License，保留所有权利。使用本代码需明确标注来源，禁止闭源商业使用。

🔄 持续更新中

