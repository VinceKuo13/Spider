# 政府采购公告爬虫使用说明

这个项目已经把 7 个省份的爬虫脚本放在同一个文件夹里。日常使用时，**只需要修改 `config.yaml` 里的关键词和时间**，然后运行对应省份的 `.py` 文件即可。

\---

## 1\. 文件说明

项目主文件夹中主要有这些文件：

```text
Spider/
├─ config.yaml          # 只改这个：关键词、开始日期、结束日期
├─ beijing.py           # 北京爬虫
├─ shanxi.py            # 山西爬虫
├─ henan.py             # 河南爬虫，需要人工输入验证码
├─ hebei.py             # 河北爬虫
├─ shandong.py          # 山东爬虫
├─ tianjin.py           # 天津爬虫
├─ xinjiang.py          # 新疆爬虫，使用后台浏览器
├─ common\_config.py     # 自动读取 config.yaml，不需要修改
├─ requirements.txt     # 依赖列表
├─ setup\_codespaces.sh  # GitHub Codespaces 环境安装脚本
└─ setup\_windows.bat    # Windows 环境安装脚本
```

\---

## 2\. 修改关键词和时间

打开 `config.yaml`，只需要改这三处：

```yaml
keywords:
  - 营商环境
  # - 政务服务
  # - 数字政府

start\_date: "2025-01-01"
end\_date: ""
```

### 2.1 修改关键词

只爬一个关键词：

```yaml
keywords:
  - 营商环境
```

爬多个关键词：

```yaml
keywords:
  - 营商环境
  - 政务服务
  - 数字政府
```

注意：每个关键词前面要有 `-`，并且要和后面的文字之间保留一个空格。

### 2.2 修改时间

例如爬取 2025-01-01 到 2026-05-09：

```yaml
start\_date: "2025-01-01"
end\_date: "2026-05-09"
```

如果 `end\_date` 留空：

```yaml
end\_date: ""
```

表示自动爬到运行当天。

\---

## 3\. 选择爬取哪个省

不需要在 `config.yaml` 里选择省份。想爬哪个省，就在终端运行哪个省的脚本。

例如：

```bash
python beijing.py
```

```bash
python tianjin.py
```

```bash
python shandong.py
```

全部脚本名称如下：

|省份|运行命令|
|-|-|
|北京|`python beijing.py`|
|山西|`python shanxi.py`|
|河南|`python henan.py`|
|河北|`python hebei.py`|
|山东|`python shandong.py`|
|天津|`python tianjin.py`|
|新疆|`python xinjiang.py`|

\---

## 4\. 在 GitHub Codespaces 上运行

### 第一步：进入项目文件夹

在 Codespaces 终端输入：

```bash
cd Spider
```

如果终端已经在 `Spider` 文件夹里，就不需要再执行这一步。

可以用下面命令确认当前位置：

```bash
pwd
```

看到路径最后是 `Spider` 即可。

### 第二步：安装环境

第一次运行前，输入：

```bash
bash setup\_codespaces.sh
```

这个命令会自动安装需要的依赖。新疆爬虫需要浏览器内核，脚本里也会自动安装。

### 第三步：修改配置

打开左侧文件列表中的 `config.yaml`，修改关键词和时间，保存文件。

### 第四步：运行省份脚本

例如运行天津：

```bash
python tianjin.py
```

运行山东：

```bash
python shandong.py
```

运行新疆：

```bash
python xinjiang.py
```

\---

## 5\. 在 Windows 本地运行

### 第一步：打开终端

在项目文件夹空白处右键，选择“在终端中打开”或“打开 PowerShell”。

### 第二步：进入项目目录

如果压缩包解压后文件夹叫 `Spider`，可以输入：

```powershell
cd Spider
```

### 第三步：安装依赖

第一次运行前，双击：

```text
setup\_windows.bat
```

也可以在终端输入：

```powershell
pip install -r requirements.txt
python -m playwright install chromium
```

### 第四步：运行省份脚本

例如：

```powershell
python tianjin.py
```

\---

## 6\. 虚拟环境说明

如果电脑已经安装了 Anaconda 或 Miniconda，建议使用单独环境。

创建环境：

```bash
conda create -n spider python=3.10 -y
```

激活环境：

```bash
conda activate spider
```

安装依赖：

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

以后每次运行前，只需要先激活环境：

```bash
conda activate spider
```

再运行：

```bash
python tianjin.py
```

\---

## 7\. 输出文件在哪里

每个省份会生成自己的输出目录，例如：

```text
outputs\_tianjin/
outputs\_shandong/
outputs\_hebei/
outputs\_henan\_auto/
outputs\_xinjiang/
```

Excel 文件就在对应的输出目录里。

\---

## 8\. 特殊说明

### 8.1 河南需要人工输入验证码

运行 `python henan.py` 后，程序会把验证码图片保存到：

```text
outputs\_henan\_auto/captcha/
```

打开最新的验证码图片，把验证码输入到终端里即可。

### 8.2 新疆使用后台浏览器

新疆站点有反爬校验，所以 `xinjiang.py` 使用 Playwright 后台浏览器。第一次运行前必须安装浏览器内核：

```bash
python -m playwright install chromium
```

在 Codespaces 中建议执行：

```bash
python -m playwright install --with-deps chromium
```

本项目里的 `setup\_codespaces.sh` 已经包含这一步。

\---

## 9\. 常见问题

### 问题 1：提示找不到某个 Python 包

例如：

```text
ModuleNotFoundError: No module named 'xxx'
```

解决：

```bash
pip install -r requirements.txt
```

### 问题 2：新疆提示浏览器不存在

解决：

```bash
python -m playwright install chromium
```

### 问题 3：没有爬到数据

请检查：

1. `config.yaml` 的关键词是否正确；
2. 日期范围是否正确；
3. 网站本身是否能打开；
4. 是否已经抓过同样的数据，部分脚本会跳过已访问链接。

可以先换一个宽泛关键词测试，例如：

```yaml
keywords:
  - 环境
```

\---

## 10\. 最常用操作总结

日常只做三步：

```text
1. 修改 config.yaml
2. 打开终端进入 Spider 文件夹
3. 输入 python 省份.py
```

例如：

```bash
python tianjin.py
```

