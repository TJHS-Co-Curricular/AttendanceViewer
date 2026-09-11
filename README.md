# 出缺席记录查看网站 (Attendance Viewer)

`app.py` 是一个本地小网站，会读取 `Result/` 文件夹底下每一个日期资料夹（例如
`01.08.2026-(例常32)出缺席表 (File responses)`），解析该资料夹内每个学会/团体
上传的 `.xlsx` 缺席记录档，并把 **序 / 姓名 / 班级 / 学号 / 缺席情况 / 备注**
整理成一个网页表格显示出来。

有两种使用方式：**A) 直接用 Python 运行**，或 **B) 打包成一个不需要 Python
的 .exe**，两者显示的内容完全一样。

---

## 项目文件结构

```
app.py                 网站逻辑（读取 / 解析 .xlsx、路由）
templates/              网页 HTML 模板
  base.html
  index.html
  folder.html
  grouped.html
static/                 样式 / 前端脚本 / 图示
  style.css
  app.js
  favicon.svg
requirements.txt        Python 依赖清单
build_exe.bat           一次性打包成 .exe（见下方「方式 B」）
README_app.md           本说明文件
Result/                 你的资料（日期命名的资料夹，內含各学会上传的 .xlsx）
```

运行或打包时，请保持以上文件/文件夹彼此相对位置不变（`app.py`、
`templates/`、`static/`、`requirements.txt`、`build_exe.bat` 都在同一层）。
`build_exe.bat` 打包时会自动把 `templates/` 和 `static/` 一起封装进
`AttendanceViewer.exe`，所以打包完成后，只有 `.exe` 本身和 `Result/`
资料夹需要放在一起——不需要额外拷贝 `templates/` / `static/`。

---

## 方式 A：直接用 Python 运行

```bash
pip install -r requirements.txt
python app.py
```

浏览器会自动打开 <http://127.0.0.1:5000>。

---

## 方式 B：打包成可携带的 .exe（不需要 Python）

> **重要说明**：一个真正的 Windows `.exe` 只能在 **Windows 电脑上打包**，
> 无法在网页/云端环境里"跨平台"生成——这是 PyInstaller（打包工具）本身
> 的限制，不是本程序的限制。但这个打包步骤**只需要做一次**：
>
> 1. 在**任何一台装有 Python 的 Windows 电脑**上，双击 `build_exe.bat`。
> 2. 它会自动安装打包所需的工具、把 `app.py` 连同 `templates/` 与
>    `static/` 一起打包成一个独立的 `AttendanceViewer.exe`，并且清理掉
>    打包过程中产生的临时文件。
> 3. 打包完成后，把这个 `.exe` 连同一份 `Result` 文件夹拷贝到**任何一台
>    Windows 电脑**（不用装 Python、不用装任何东西），双击就能用。
>
> （`build_exe.bat` 内容刻意只用英文字母，避免中文批处理档在不同电脑的
> 「代码页」设定下出现乱码或异常——这是纯技术考量，跟程序功能无关。
> 网页界面本身仍然完全是中文显示。）

### 步骤

1. 确保这台电脑装有 Python（打包**这一次**需要用到）。
   若还没装，到 <https://www.python.org/downloads/> 下载安装，安装时
   记得勾选 **"Add python.exe to PATH"**。
2. 双击 `build_exe.bat`（与 `app.py`、`templates/`、`static/`、
   `requirements.txt` 放在同一个文件夹）。
3. 等待它跑完（大约 1-3 分钟），完成后同一个文件夹里会多出
   `AttendanceViewer.exe`（`templates/` 和 `static/` 的内容已经打包在
   里面了，不需要额外拷贝）。
4. 之后，把这个 `.exe` 和你的 `Result` 文件夹放在同一层目录，双击 `.exe`
   即可使用——**这台电脑或任何其他电脑都不需要安装 Python**。

### 分享给别人使用

打包完成后，只需要把以下两样东西一起拷贝给对方（例如整个塞进一个
压缩包 / U 盘）：

```
AttendanceViewer.exe
Result/                 ← 你的资料夹
```

对方直接双击 `.exe`，浏览器就会自动打开，完全不需要安装 Python。

---

## 功能说明

- 首页列出 `Result/` 底下所有日期资料夹。
- 点击任一日期，会显示当天所有学会（依 A01 / A02 / B01... 代码顺序排列）
  上传的缺席记录，並可以：
  - 用右上角搜寻框依姓名 / 班级 / 学号 / 备注过滤
  - 点击「下载 CSV」把当天所有缺席记录汇出成一份 CSV（可用 Excel 打开）
  - 点击「依缺席情况分组名单 ↗」，在新分页打开一份按 **缺席情况类型**
    （旷课 / 病假 / 事假 / 特别事假 / 公假…）分组、再按学会代码排序、
    每 10 人一行的学号清单，方便直接复制或下载 `.txt`
    （已自动略过「内部公假」「迟到」「早退」这三种类型，这三类仍会
    正常显示在主表格与 CSV 里，只是不出现在这份分组清单中）。
- 若某学会勾选「全勤」且没有列出任何缺席学生，会显示绿色
  「全勤 / 无缺席」标记。
- 若某个档案格式跟预期不符、无法解析，网页上会用红字标注「解析出错」，
  不会让整个网页当掉。
- 每次点击一个日期资料夹，网页都会强制重新读取最新资料，不会显示浏览器
  缓存的旧内容（但已经解析过、档案没有变动的 `.xlsx` 仍会用内部缓存
  加速，不影响正确性）。
- 如果你的 `Result` 文件夹跟 `app.py` / `.exe` 不在同一层，也可以指定路径：
  ```
  python app.py "D:\某个地方\Result"
  AttendanceViewer.exe "D:\某个地方\Result"
  ```
- 默认情况下，终端窗口只显示开机横幅（网址等），不会为每次点击都印出一行
  记录，保持画面干净。如果需要看到每一次请求的详细记录（除错用），用
  `--dev` 参数启动即可：
  ```
  python app.py --dev
  AttendanceViewer.exe --dev
  ```
  （也可以设定环境变量 `ATTENDANCE_VIEWER_DEBUG=1` 达到同样效果。）
