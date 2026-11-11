# 求职投递同步

本地小工具。用本机 Edge 打开 Boss直聘、猎聘、智联招聘，读投递列表，合并进你自己的 Excel。

同一平台、同一公司、同一职位只保留一行。状态有变化才改（例如「已发简历」变成「不合适」），不会把新状态改回旧状态。

手机号、Excel 路径、Edge 路径都只写在你自己的 `config.json` 里。程序里没有这些默认值。没有 `config.json` 时，程序只生成一份空的 `config.template.json` 然后停下来，不会替你填，也不会用别人的号码去登录。

## 环境

| 项目 | 本机实测 | 说明 |
| --- | --- | --- |
| 操作系统 | Windows 10 | 启动脚本是 `.bat`，按 Windows 写的 |
| Python | **3.7.5** | 官方安装包，勾选了 Tcl/Tk |
| 浏览器 | Microsoft Edge | 用系统里已经装好的 Edge，不另下浏览器 |
| 界面 | tkinter | Python 自带，`requirements.txt` 里没有 |

### 需要安装的库

只有两个，都写在 `requirements.txt` 里：

- `openpyxl`：读写 Excel
- `websocket-client`：连 Edge 的调试端口，接收网页请求

不需要安装 Playwright 或 Selenium。程序直接使用系统里已经安装的 Edge。

### Python 版本不一样时

代码按 **3.7** 写的：没有 `list[str]` 这种写法，没有海象运算符，字符串用 `%` 格式化。3.7 到 3.12 一般可以直接跑。

- **低于 3.7**：不要用。
- **3.8–3.12**：推荐。`requirements.txt` 里的版本号可以直接装。
- **3.13 及以上**：`openpyxl==3.1.3` 或 `websocket-client==1.6.1` 可能装不上。这时去掉版本号再装：

  ```bat
  .venv\Scripts\python -m pip install openpyxl websocket-client
  ```

  业务代码没有绑死这两个库的内部接口。`websocket-client` 需要保留 `create_connection(..., suppress_origin=True)`，1.6 及之后的版本都有。

- 安装 Python 时必须带 **Tcl/Tk**（python.org 安装器里默认勾着）。精简版、嵌入版经常没有 tkinter，一运行就会报 `No module named 'tkinter'`。
- 公司网络如果 pip 报 SSL 证书错误，换国内镜像：

  ```bat
  .venv\Scripts\python -m pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
  ```

## 下载后怎么用

仓库里没有任何人的手机号、登录状态或投递记录。下载 ZIP 或 `git clone` 之后，按这个顺序做：

1. 安装 Python 3.7 到 3.12（安装时勾选 **tcl/tk and IDLE**）和 Microsoft Edge。
2. 进入解压出来的目录，打开命令行：

```bat
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

3. 把 `config.template.json` 复制为 `config.json`，用记事本填你自己的手机号、Edge 路径和下面这份表格的路径。
4. 把 `投递记录模板.xlsx` 复制到别的文件夹，例如「文档」。`config.json` 的 `excel` 填这个副本的完整路径。不要直接用程序目录里的模板，以免以后更新代码时和你的投递数据混在一起。
5. 关掉这份 Excel，双击 `同步投递记录.bat`。

模板里已有「投递记录」「统计」「使用说明」三张表，表头、下拉选项和重复投递的颜色都设好了。使用简历的下拉只有「默认简历」和「定制版」，要改成你自己的简历名称时，在 Excel 里选中 G 列，用「数据验证」修改。

下面是同一套步骤的详细说明，以及运行到一半时每个弹窗该怎么处理。

## 第一次安装

1. 安装 Python 3.7 到 3.12，安装时保持勾选 **tcl/tk and IDLE**。再确认电脑上有 Microsoft Edge。
2. 在本目录打开命令行，执行：

```bat
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

3. 准备 `config.json`。目录里如果已经有一份填好的，就只改成你自己的内容。如果没有：
   - 双击一次 `同步投递记录.bat`，程序会弹出提示并生成 `config.template.json`，然后停住。
   - 或者直接使用目录里的 `config.template.json`。
   - 把 `config.template.json` **复制**一份，改名为 `config.json`。不要改模板原件，也不要把别人的 `config.json` 拷过来。
4. 用记事本打开 `config.json`，按每个 `_` 开头的说明填写。以 `_` 开头的行可以留着，程序不读它们。必须由你自己填写的是：

| 字段 | 你要填什么 |
| --- | --- |
| `phone` | 你自己的 11 位手机号，不要加 +86 或空格 |
| `excel` | 你复制出来的 `投递记录模板.xlsx` 的完整路径。文件必须已存在，而且有「投递记录」这个工作表 |
| `edge` | 你电脑上 `msedge.exe` 的完整路径。开始菜单里对 Microsoft Edge 右键，选「打开文件所在的位置」 |
| `port` | 一个整数。本机没被占用即可，可以写 `9333` |
| `max_pages` | 每个列表最多翻多少页，正整数，例如 `30` |
| `debug_dump` | `true` 或 `false`，不能留空 |
| `boss_contact_include` | Boss「沟通过」要保留的职位关键词，不需要就写 `[]` |
| `boss_contact_exclude` | Boss「沟通过」要丢掉的职位关键词，不需要就写 `[]` |

`boss_contact_include` 留空时，Boss「沟通过」会全部导入。Boss「已投递」、猎聘、智联不受这两个数组影响。

5. 保存 `config.json`。哪一项是空的，或者路径不对，程序都会列出问题并停住，不会自己补上。

## 每次运行前

1. 关掉正在打开这份 Excel 的窗口。开着的话，最后写不进去。
2. 手机放在手边。第一次使用，或登录过期时，三个网站可能各发一条短信验证码。
3. 不要同时再开一个「同步投递记录」。它会占用 `config.json` 里的 `port`。

双击 `同步投递记录.bat`。会留下一个黑色命令行窗口，同步过程的文字打在里面。先不要关它。结束时这个窗口会停在「请按任意键继续」，看完结果再按键关掉。

## 运行过程中你要做什么

程序会自己打开一个 Edge 窗口。这是专用窗口，登录状态记在本目录的 `edge_profile` 里。不要关掉它，也不要在它正在翻页时自己点别的链接。你只在弹出对话框时动手。

### 1. 选择网站

弹出「求职投递同步」。三个网站默认都勾着。

- 不想同步某一个，就取消勾选。
- 「预览模式」勾上后只抓取、不改 Excel。第一次拿不准时可以先勾上，看黑色窗口里的「预览」行。
- 点「开始同步」。点窗口右上角的叉则直接退出，不打开网站。

### 2. 已经登录过

黑色窗口出现「已是登录状态」时，你什么都不用做。程序会自己翻「沟通过 / 已投递 / 全部状态 / 投递成功 / 不合适」。这一段可能要一两分钟，Edge 窗口会自己跳页，不要去点它。

三个网站是一个接一个做的。前一个抓完才会打开下一个。

### 3. 需要登录：等短信，并输入验证码

某个网站还没登录时，窗口会写「需要登录，开始手机号验证码登录」。程序会自己：

1. 打开该网站的登录页。
2. 填入 `config.json` 里的手机号。
3. 勾选「已阅读并同意」。
4. 点击「发送验证码」。

然后弹出标题为「输入短信验证码」的框，文字类似：

> Boss直聘 已点击「发送验证码」到 你的手机号。
> 如果浏览器弹出滑块/点选验证，请先在浏览器里完成它。
> 收到短信后在下面输入验证码（点取消=我在浏览器里自己登录）

这时请你：

1. 先看 Edge 里有没有滑块、点选图片之类的验证。有的话，在 **浏览器里** 做完，不要关浏览器。
2. 看手机短信。可能要等十几秒到一分钟。这个输入框会一直停着，直到你点「确定」或「取消」。
3. 把短信里的数字填进框里，点「确定」。不要带空格。
4. 程序会把验证码填进网页并点「登录/注册」。登录成功后继续抓列表，你不用再管。

点「取消」表示这次不替你填验证码。接着会再弹一个说明框，请你在浏览器里自己登录，登完后点「确定」，程序再去检查登录是否成功。

三个网站的登录是分开的。第一次使用最多会出现三次验证码框，每次都是当前这个网站的短信。已经登录过的网站不会再要。

### 4. 安全验证或滑块

Boss 有时会整页变成「安全验证」，提示当前 IP 可能异常。猎聘或智联也可能弹出滑块。

这时程序停住，弹出说明框，例如：

> Boss直聘 出现了安全验证页面。请在浏览器里点按钮/拖滑块完成验证，完成后点「确定」继续。

你在 Edge 里把验证做完，再回到这个框点「确定」。不要点确定之前就把浏览器关了。

### 5. 自动登录没成功

可能出现下面几种框，处理方式一样：你在 Edge 里把这个网站登录完，再点「确定」。

- 「没找到手机号输入框，请在浏览器里手动登录」
- 「仍未登录。请在浏览器里手动登录」
- 「自动提交后似乎还停在登录页」

手动登录时同样要勾选协议。点「确定」之后程序会重新打开投递列表；如果这时已经登录，就接着抓。

### 6. Excel 正被打开

写入前会先把原文件复制到 `backups`。如果 Excel 还开着，会弹出：

> Excel 文件正被打开。请关闭 Excel 后点「重试」；点「取消」则另存为新文件。

关掉 Excel 后点「重试」，结果写回原文件。点「取消」则另存为同一目录下的 `原文件名_同步_时分秒.xlsx`，原文件保持打开前的内容。

### 7. 结束

全部网站跑完后弹出汇总：新增多少条、状态变成了什么。点「确定」即可。黑色窗口里有同样的文字。然后按任意键关闭黑色窗口。

打开 Excel 后可以看两处：

- 「投递记录」里的新行和变色的状态格。「不合适」是红色，「已查看」是黄色。
- 「同步日志」里每次新增和状态变化各有一行。

预览模式不会改 Excel，结果只在黑色窗口和 `logs` 目录里。

命令行也可以跑，效果和上面的对话框相同，只是不再让你选网站：

```bat
.venv\Scripts\python job_sync.py --sites boss,liepin,zhaopin
.venv\Scripts\python job_sync.py --sites liepin --preview
```

`--sites` 时结束不弹汇总框，结果在命令行里。验证码、安全验证、Excel 被占用，仍然会弹框等你操作。

## 跑完 Excel 里有什么

写入的是「投递记录」表，列顺序不变：

投递日期、平台、公司名称、职位名称、城市、薪资、使用简历、状态、下次跟进日期、联系人 / HR、职位链接、备注。

程序只填它抓得到的列，不改「使用简历」和「下次跟进日期」。原有的数据验证、条件格式、M/N 列计数公式都保留。新增行超过第 301 行时，会把同样的计数公式续写下去。

另外会建（或追加）一张「同步日志」表，记下每次新增和状态变化。

状态对应关系：

| 网站上的情况 | 写入的状态 |
| --- | --- |
| Boss「沟通过」（对方没回复，简历发不出去） | 已打招呼 |
| Boss「已投递」、猎聘「已投递」、智联「投递成功」 | 已发简历 |
| 猎聘「已查看」。接口里的原文是「筛选中」 | 已查看 |
| 智联「不合适」，或猎聘「暂不合适」 | 不合适 |
| 文案里出现「面试」 | 面试邀请 |
| 文案里出现 offer / 录用 | offer |

两个列表都有同一条时，取更靠后的状态。Boss 上既在「沟通过」又在「已投递」，记成「已发简历」。

更新规则：

- 匹配键是 **平台 + 公司 + 职位**，三个都要像才算同一条。
- 公司名会去掉「有限公司 / 股份 / 集团」等后缀，所以「海信集团」和「海信集团控股股份有限公司」能对上。
- 职位名忽略空格和全角括号。两边职位编号相同（如 `J45689`）也算同一条。短标题是长标题的开头、且不少于 6 个字时也算同一条。
- 状态只前进：已打招呼 → 已发简历 → 已查看 → 面试邀请 → 一面 → 二面 → 终面 → offer。变成「不合适」也会更新。
- 不会从「已查看」退回「已发简历」。你手工填的「一面 / offer / 已放弃」不会被网站上的旧状态盖掉。
- 状态格「不合适」标红，「已查看」标黄。其它状态不上色。备注末尾追加一句，例如 `11-11 同步：已发简历→不合适`。

不同平台的同一职位各记一行。Boss 上打过招呼、猎聘上又投过简历，是两次投递。

## 代码结构

```
同步投递记录.bat   双击入口，用本目录 .venv 里的 Python 跑 job_sync.py
job_sync.py        总控：读配置、弹窗、按网站循环、汇总后写 Excel
sites.py           三个网站的登录步骤和列表抓取
cdp.py             启动 Edge，并通过调试协议操作页面、抓接口 JSON
extract.py         从接口 JSON 里抽出公司、职位、薪资、状态、时间
excel_sync.py      和现有表格比对、去重、改状态、备份、处理文件被占用
config.json        你自己填的手机号、Excel 路径、Edge 路径、关键词。不要交给别人
config.template.json  空模板。没有 config.json 时运行会再生成一份，程序不会拿它直接跑
投递记录模板.xlsx  空的投递记录工作簿。复制到别处再使用，不要把个人投递写进仓库里的这一份
create_template.py     重新生成上面的空工作簿
.gitignore             排除 config.json、edge_profile、debug、logs、backups、.venv
edge_profile\      Edge 的独立用户目录，里面是登录 Cookie。不要提交、不要随手删
logs\              每次运行一份日志
debug\             debug_dump 打开时，每个列表一份 json / 页面文本 / 截图
backups\           每次写入前的 xlsx 副本
```

一次运行的顺序：

```
bat / 命令行
    → job_sync.main
        → 选择网站（窗口或 --sites）
        → cdp.Browser 启动 Edge（调试端口 + edge_profile）
        → 每个网站 new_tab
            → sites.Site.ensure_login
            → sites.Site.scrape_view（按标签翻页）
            → cdp 把接口响应存进 tab.captured
            → extract.walk 变成统一的记录
        → excel_sync.merge 写入「投递记录」和「同步日志」
        → 弹出结果
```

### `job_sync.py`

入口。`load_config` 只读 `config.json`。文件不存在时写出空的 `config.template.json` 并停止；缺项、空手机号、路径不存在也会停止。代码里没有手机号、表格路径或 Edge 路径的默认值。

`UI` 是三个 tkinter 对话框：开始时选网站和预览模式；登录时要验证码；Excel 被占用时问要不要重试。命令行加了 `--sites` 就不再弹开始窗口和结束提示，方便重复跑。

`main` 不自己解析网页。它只按 `sites.SITES` 建标签页、调用 `Site.run()`，把各站记录收齐后交给 `excel_sync.merge`。某一站出错只记日志，不影响另外两站。

### `cdp.py`

自己写的一个很小的 Chrome DevTools 客户端，没有用 Playwright，也没有用 Selenium。

原因：猎聘会检测常见自动化库。用 Playwright 打开时，页面会被脚本改成 `about:blank`。现在的做法是让 Edge 以普通进程启动，只带 `--remote-debugging-port`，再连上去。连上之后**不发送** `Runtime.enable` / `Page.enable`，这两个也会触发同样的检测。只开 `Network.enable` 来收接口数据。

- `Browser.start`：端口上已经有 Edge 就复用，否则用 `edge_profile` 新开一个。
- `Tab.goto`：用页面里的 `location.href` 跳转，不用 Playwright 的 `page.goto`。
- `Tab.click` / `Tab.type`：先算元素中心坐标，再发 `Input.dispatchMouseEvent` 和 `Input.insertText`。这是浏览器认为的真实点击，直接改 DOM 点按钮，网站经常不认。
- `Tab.captured`：接口返回的 JSON。抓取逻辑只读这里，不抠 HTML。

### `sites.py`

`SITES` 是三站的差异：登录框选择器、列表地址、要看的标签、哪一个接口算数。

| 网站 | 页面 | 读取的接口 |
| --- | --- | --- |
| Boss直聘 | 「沟通过」逐页打开 `tag=5` | `/wapi/zprelation/interaction/geekGetJob?tag=5` |
| Boss直聘 | 「已投递」 | `/wapi/zprelation/resume/geekDeliverList` |
| 猎聘 | 「全部状态」 | `api/com.liepin.capply.pc.apply-list` |
| 智联招聘 | 「投递成功」「不合适」 | `/c/i/schedule/feedback?status=send` 和 `status=unsuitable` |

Boss 侧边栏每页都会顺带请求「感兴趣」（`tag=2`），所以 `api_include` 必须带 `tag=5`，否则会把没投过的推荐职位算进来。

`Site.ensure_login` 发现还在登录页或验证页时，走手机号流程；验证码不自动填。自动登录失败两次后，提示你在浏览器里手动登录。

`scrape_view` 有两种翻页。Boss 的页码在网址里，直接打开 `page=1,2,3...`，某一页没有新记录就停。猎聘和智联是点「下一页」或滚到底。`map_status` 把网站原文收成表格里那一套状态；猎聘这条视图开了 `use_text`，所以「筛选中」会变成「已查看」，而不是一律写成「已发简历」。

Boss「沟通过」额外做 `title_filter`，用 `config.json` 里的两组词。

### `extract.py`

`walk` 递归扫接口 JSON。一个字典里同时能找到「职位名」和「公司名」就当成一条记录，不再往下拆，避免把公司对象再识别成一条假职位。

字段名在各站不一样，所以按候选键找：`jobName` / `title`、`brandName` / `compName`、`salaryDesc`、`applyStautsDesc`（猎聘接口里就是这个拼写）、`applyTime` / `happenTime`。时间戳和 `2026-11-11 10:18` 这种字符串都会转成日期。

### `excel_sync.py`

`merge` 打开工作簿，从第 2 行扫到公司列为空。对每条新记录：

1. 同一平台下，公司和职位都匹配 → 补上原来空着的城市、薪资、链接、HR、日期；状态满足 `should_update` 才改。
2. 匹配不上 → 追加一行，备注写「自动同步新增」。

`save` 先复制到 `backups\`，再保存。遇到 `PermissionError` 就调用上面传入的「重试」对话框。

## 换到另一台电脑

要复制的：

- `job_sync.py`、`sites.py`、`cdp.py`、`extract.py`、`excel_sync.py`
- `requirements.txt`、`README.md`、`同步投递记录.bat`、`config.template.json`

不要复制：

- `.venv\`：里面的路径和 Python 版本绑着旧电脑，到新电脑会坏。按上面的步骤重新建。
- `config.json`：里面是某一个人的手机号和表格路径。换人或换电脑时不要把这份文件一起拷走，让对方从 `config.template.json` 填自己的。
- `__pycache__\`、`logs\`、`debug\`、`backups\`：可有可无，不参与运行。

看情况复制：

- `edge_profile\`：复制过去可以少登录一次。这个目录可能很大，而且和本机 Edge 版本绑在一起；新电脑登录不上、或一打开就报用户数据损坏时，删掉它再运行，程序会新建，然后重新收一次验证码。
- 投递记录工作簿不在仓库里。新电脑上把原来的工作簿拷过去，或从 `投递记录模板.xlsx` 复制一份空表，再把路径写进新的 `config.json`。程序不会自己新建工作簿。

新电脑上还要做：

1. 安装 Python 3.7–3.12（带 Tcl/Tk）和 Microsoft Edge。
2. 按「第一次安装」建虚拟环境并 `pip install -r requirements.txt`。
3. 从 `config.template.json` 复制出新的 `config.json`，填这台电脑上的手机号、Excel 路径和 `msedge.exe` 路径。`port`、`max_pages`、`debug_dump` 也要自己填，程序不带默认值。
4. 确认没有别的程序占用你填的 `port`。上次异常退出留下的 Edge 会占着这个端口，任务管理器里结束对应的 `msedge.exe`，或在 `config.json` 里换一个端口。
5. 双击 `同步投递记录.bat`。第一次会要验证码；三个网站各自登录一次即可。

`同步投递记录.bat` 写死了 `.venv\Scripts\python.exe`。虚拟环境如果不用这个名字，就要改 bat，或者始终按文档创建 `.venv`。

## 抓到 0 条时

网站改版后选择器或接口路径会变。`debug_dump` 为 true 时，`debug\日期时间\` 下每个列表有三份文件：

- `.json`：当时收到的接口响应
- `.txt`：页面上的文字
- `.png`：截图

日志里会出现「没抓到数据」。把对应的 json 和 png 留着，对照 `sites.py` 里该站的 `api_include`、`tab_text`、登录选择器改即可，不必改 `excel_sync.py`。
