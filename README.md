# onenote-plugin

一个 **多宿主 OneNote 插件**：同一份核心同时支持 **Claude Code（stdio MCP）** 与 **DSH（原生 Cordis）**，直接读写本机 **Microsoft OneNote 桌面版**（COM 自动化），操作的是真实本地笔记本（含 `.one` 分区文件），不是云端的 Graph 副本。

- 工具：`onenote_list` / `onenote_read` / `onenote_search` / `onenote_create` / `onenote_update` / `onenote_append` / `onenote_edit` / `onenote_insert_image` / `onenote_normalize_tables` / `onenote_publish` / `onenote_delete` / `onenote_create_section` / `onenote_rename_section`
- 依赖：仅 `@modelcontextprotocol/server`；运行时还需要 Windows PowerShell 5.1（`powershell.exe`）+ 本机 OneNote
- 架构：每次工具调用 → 写 `req.json` → spawn `powershell.exe` 跑 `scripts/onenote-bridge.ps1`（COM）→ 读 `res.json`（§13 A 方案）

## 安装

```powershell
cd <项目目录>
npm install
```

## 注册到 Claude Code

```powershell
claude mcp add onenote -- node <项目目录>\server.js
```

带配置（可选，环境变量用 `-e` 传）：

```powershell
claude mcp add onenote -e ONENOTE_EXTRASECTIONS="[{\"name\":\"快速笔记\",\"path\":\"<你的路径>\\快速笔记.one\"}]" -e ONENOTE_PAGECOLOR="#C7EDCC" -- node <项目目录>\server.js
```

> 独立分区的 `name` 只是一个**别名**（用于 `sectionName` 寻址），建议取唯一名字：
> 这台机器的笔记本 `Technology/Virtualization/KVM` 里已有一个叫 `KVM` 的分区，
> 无别名时按 `sectionName: "KVM"` 会命中那个笔记本分区。
> 独立分区本身不允许 `UpdateHierarchy` 改名（`0x8004200B SECTION_READ_ONLY`），所以别名只能写在配置里。

## 作为 Claude Code 插件安装（发布版）

仓库同时是一个 Claude Code 插件：内含本 MCP server、`skills/onenote` 技能和
`.mcp.json`，用 `${CLAUDE_PLUGIN_ROOT}` 定位 `server.js`，不依赖机器上的绝对路径。

```text
/plugin marketplace add AlawnWalker/onenote-plugin
/plugin install onenote-plugin@onenote-plugin
```

- 首次启动由 `scripts/start.mjs` 在插件目录自动执行 `npm install --omit=dev`；
  之后的会话直接启动。首次安装依赖需要联网。
- 插件是 **Windows + OneNote 桌面版专用**。
- 如果本机已经用 `claude mcp add onenote ...` 注册过同名 server，安装插件前先
  `claude mcp remove onenote -s user`，避免两套同名工具。
- 独立 `.one` 分区用 `~/.cc-onenote.json` 配置（见下），不需要改插件文件。
- 在插件仓库根目录直接运行 Claude Code 时，Claude Code 会把根目录的 `.mcp.json`
  同时当作项目级 MCP 配置，显示一个等待批准的 `onenote` 条目；这是开发时的副作用，
  正常从其他目录使用插件不需要批准它。

## 作为 DSH 插件安装（原生 Cordis）

同一仓库 / 同一个 npm 包也包含 DSH 原生入口（`dsh/index.js` + `cordis.patch.yml`）。
安装脚本会同步源码到 `DSH_HOME`、登记 profile 依赖与 bundle，并移除旧的 `dsh-onenote`
条目（旧目录保留作回滚）：

```powershell
# 先看计划（不写任何文件）
powershell -ExecutionPolicy Bypass -File .\install.ps1 -DryRun

# 真正安装（DSH 关闭或在运行中都行；装完必须重启桌面端）
powershell -ExecutionPolicy Bypass -File .\install.ps1 `
  -DshCli "<DSH 安装目录>\resources\runtime\cli\bin\dsh.cmd" `
  -Registry http://172.27.30.100:8765/
```

- `DSH_HOME` 决定装进哪个 home（桌面端读取它）；本机是
  `D:\Leon\Deepseek\Workpsace`，不要装到 `%USERPROFILE%\.dsh` 的旧 home。
- 重启后工具名是 `onenote_*`（DSH 原生），不是 `mcp__...` 前缀。
- `install.ps1 -ExtraSectionName/-ExtraSectionPath` 只写入**已安装包内**的
  `cordis.patch.yml`；如果 profile 里存在同 `id: dsh-onenote` 的覆盖条目，
  **覆盖会整块替换 config**，此时应把 `extraSections` 写进 profile 覆盖
  （脚本会给出警告）。

## 配置（环境变量）

| env | 默认 | 说明 |
|---|---|---|
| `ONENOTE_POWERSHELL` | `powershell.exe` | 必须是 Windows PowerShell 5.1（不是 pwsh 7） |
| `ONENOTE_TIMEOUT_MS` | `180000` | 单次调用超时；冷启动会拉起 OneNote，第一次可能 1–3 分钟 |
| `ONENOTE_PAGECOLOR` | `#C7EDCC`（豆沙绿） | `#RRGGBB` / `automatic` / `none`；设空字符串 `""` 跟随 OneNote 默认 |
| `ONENOTE_PAGESIZE` | 空（Automatic） | JSON 字符串，如 `{"width":1190,"height":842,"landscape":true}`，或 `"automatic"` |
| `ONENOTE_CODE_PREFIX` | `│ ` | 代码块每行的可视前缀；设 `""` 只去掉竖线，读回仍按 Consolas/monospace 字体识别为代码 |
| `ONENOTE_ALIGN_COMMENT_COLUMN` | `0` | `0` 不处理 / `auto` 对齐到最长代码+2 / `N` 固定列（按字节算） |
| `ONENOTE_ALIGN_COMMENT_MAX` | `80` | `auto` 对齐的列上限 |
| `ONENOTE_BULLET_VALUE` | `1` | OneNote 项目符号字形索引（0 大黑圆、1 实心点 U+2022、3 空心圆、7 菱形） |
| `ONENOTE_BULLET_VALUES` | `[]` | JSON 数组，按层级指定字形，如 `[20, 2]`；最后一项重复 |
| `ONENOTE_HEADING_DELTA` | `[2, 1]` | 标题字号相对正文的增量 `[h2, h3]`（也接受 `{"h2":2,"h3":1}`）；旧版页面写的 `[4, 2]` 仍能正确读回 |
| `ONENOTE_BULLET_FONT_SIZE` | `body` | 项目符号字号：`body` 跟随正文字号，或填具体 pt（如 `12`） |
| `ONENOTE_NEST_UNDER_HEADINGS` | 空（关） | `1`/`true` 把标题后的内容按层级缩进（DSH 原生索引） |
| `ONENOTE_INDENTED_CHILDREN` | 空（关） | `1`/`true` 把列表项后的缩进行当作子段落 |
| `ONENOTE_EXTRASECTIONS` | `[]` | JSON 数组。**独立分区必须按 `path` 记**（ID 每次重启会变），如 `[{"name":"快速笔记","path":"<你的路径>\\快速笔记.one"}]` |
| `ONENOTE_BACKUP_DIR` | 空（`%TEMP%\onenote-plugin-backups`） | `onenote_update` / `onenote_delete` / `onenote_insert_image` / `onenote_normalize_tables` 动页面之前，先把整页 XML（含图片 base64）写到这个目录；返回里带 `backupPath` |
| `ONENOTE_PRESERVE_MEDIA` | `1` | `onenote_update` 重写 Outline 时把图片/墨迹/附件搬回去；`0` 明确丢图（会写 warning） |
| `ONENOTE_IMAGE_MAX_WIDTH` | `600` | markdown 图片未写宽度时的默认宽度（pt），超宽图按此等比缩小 |
| `ONENOTE_TREE_STYLE` | `paragraph` | ```` ```tree ```` 围栏的渲染方式：`paragraph` = 正文段落（微软雅黑/正文字号，无段距）；`mono` = 旧的等宽代码块（Consolas） |
| `ONENOTE_TREE_INDENT` | `bar` | 每级缩进用哪种字符：`bar` = 5×NBSP + U+2004（实测 20.88pt/级，最接近 `├─ ` 的 20.51pt）；`nbsp` = 6×NBSP（21.29pt）；`ideographic` = 1×U+3000（12pt，旧行为，会逐级左漂） |
| `ONENOTE_TREE_INDENT_STEP` | `3` | 源 markdown 里一个层级对应多少个半角空格（`├─ ` 宽 3 列）；用于从源缩进推导层级 |

### 用户配置文件（插件启动器）

插件通过 `scripts/start.mjs` 启动；它在启动前读取 `~/.cc-onenote.json`
（可用 `CC_ONENOTE_CONFIG` 覆盖路径），把键值写入环境变量，已存在的环境变量优先。
数组/对象会序列化为 JSON，正好对应 `ONENOTE_EXTRASECTIONS` 这类配置：

```json
{
  "ONENOTE_EXTRASECTIONS": [
    { "name": "GICv3Virtualization", "path": "Z:\\...\\GICv3 虚拟化.one" }
  ],
  "ONENOTE_PAGECOLOR": "#C7EDCC"
}
```

手动 `claude mcp add ... server.js` 的启动方式不经过这个启动器，仍需用 `-e` 传环境变量。

### DSH profile 配置覆盖（整块替换）

在 `DSH_HOME\profiles\desktop\cordis.patch.yml` 里按 `id: dsh-onenote` 覆盖插件配置。
**注意：覆盖是整块替换，不是深合并**——只写一个键会让其余键回落到插件默认值
（页面底色、代码前缀、符号、嵌套、独立分区都会变）。迁移时请把完整块抄过去：

```yaml
- id: dsh-onenote
  name: onenote-plugin/dsh
  config:
    codePrefix: ''
    alignCommentColumn: 'auto'
    alignCommentMax: 80
    pageColor: '#C7EDCC'
    bulletValues: [20, 2]
    headingDelta: [2, 1]
    bulletFontSize: 'body'
    nestUnderHeadings: true
    indentedChildren: true
    extraSections:
      - name: '快速笔记'
        path: 'C:\Users\<you>\Documents\OneNote Notebooks\快速笔记.one'
```

## 工具

| 工具 | 关键参数 | 说明 |
|---|---|---|
| `onenote_list` | `notebook?` `section?` `includePages?` `includeRecycleBin?` `maxSections?` `maxPages?` | 笔记本 → 分区 → 页面树，返回所有 ID |
| `onenote_read` | `pageId` 或 `sectionId`/`sectionName` + `pageTitle`、`maxBlocks?`、`withIds?` | 读页面，转 Markdown；`withIds=true` 在模型可见输出末尾附 `blocks (objectId)` 列表（CC 与 DSH 均生效，供 `onenote_edit` 用）。返回 `images`（类型/宽高/objectId）、`media` 计数和 `revision`（页面 XML 的 SHA-1 前 16 位）——判断"页面到底变没变"要看 `revision`，别信 `lastModified`（见限制 13） |
| `onenote_search` | `query` `scopeId?` `limit?` `includeContent?` `maxScan?` | 自扫层级；标题子串（快），`includeContent=true` 再扫正文（约 50ms/页） |
| `onenote_create` | `sectionId`/`sectionName`、`title`、`markdown` | 建页；标题也可由 markdown 的 `# 标题` 提供 |
| `onenote_update` | 页面定位 + `title?` + `markdown` + `preserveMedia?` + `backup?` | 改写已有页：删掉原文本 Outline 后写新内容；**图片/墨迹/附件自动原样搬回**（按原锚点段落就近放，锚点没了就放末尾），并把旧页 XML 备份后返回 `backupPath` |
| `onenote_append` | 页面定位 + `markdown` | 页面末尾追加一个 Outline |
| `onenote_edit` | 页面定位 + `mode` + `objectId`/`find` + 模式参数 | **原地改一个 block**：`replaceText`/`replaceBlock`/`insertAfter`/`deleteBlock`；`find` 默认子串匹配，`exact=true` 要求整块文本相等；`deleteBlock` 会连带删除嵌套子块 |
| `onenote_insert_image` | 页面定位 + `path` + `width?` + `objectId?`/`find?` + `position?` | 往**已有页面**插一张本地图片：锚点块之后/之前或页尾；只重写锚点所在的 Outline，其余块不动；插入前备份页面。文本锚点连续插入时按调用顺序排在已有图之后，`objectId` 锚点则紧贴该块；锚点块带子块时 `after` 落在整个子块之后（0.7.3） |
| `onenote_normalize_tables` | 页面定位 + `dryRun?` + `backup?` | 修复 0.7.4 之前写入的表格：把单元格段落字号钉到正文基准（同时改 OE 的 `style` 和单元格文本里的内联 `span`）；只重写含表格的 Outline，其余内容原样保留；默认先备份整页 XML；`dryRun=true` 只报告 `tables`/`cellsScanned`/`cellsChanged`/`outlinesUpdated`，不写页面（0.7.5） |
| `onenote_publish` | 页面定位 + `path` + `format?` | 导出 `.one`/`.pdf`/`.docx`/`.xps`/`.mht`/`.html` 等；目标扩展名会按格式归一化（`format=html` 写 `.htm`） |
| `onenote_delete` | 页面定位 + `permanent?` + `backup?` | 删页，默认进回收站；删之前把整页 XML（含图片）备份，返回 `backupPath` |
| `onenote_create_section` | `path`（.one 绝对路径）、`keepBlankPage?` | 建/打开独立分区；新建时自动生成的空白页会删掉，`keepBlankPage=true` 保留 |
| `onenote_rename_section` | `sectionId` 或 `path`、`name` | 改分区名；**独立分区只读**（`0x8004200B`），笔记本内分区可改 |

页面定位一律：`pageId` 优先；否则 `sectionId`/`sectionName` + `pageTitle`（先精确后子串）；只给 `pageTitle` 全库查找。

## Markdown 子集与格式约定

OneNote 2024 的 `UpdatePageContent` 拒绝 `one:Style`/`QuickStyles`/`quickStyleIndex` 这类整段样式，但**逐段的 `style` 属性 + `spaceBefore`/`spaceAfter` 可以写入**，所以标题用真实字号表达；换行（`<br/>`）仍不支持，其余层级用符号：

| Markdown | 写入 OneNote | 读回 |
|---|---|---|
| `# 标题` | 页面标题（真 Title） | `# 标题` |
| `## 标题` | 真字号段落：正文 + `headingDelta[0]`（默认 +2pt），段前 8pt / 段后 5pt | `## 标题` |
| `### 标题` | 真字号段落：正文 + `headingDelta[1]`（默认 +1pt），段前 5pt / 段后 5pt | `### 标题` |
| `> 引用` | `▎ 引用` 段落 | `> 引用` |
| ``` 代码块 | 每行一个代码段落（正文尺寸 + Consolas；`codePrefix` 非空时再加强制前缀） | 代码块（语言标记丢失；`codePrefix=''` 也能靠字体识别） |
| `- 项` / 缩进 | 真项目符号列表 | 缩进还原 |
| `1. 项` | 项目符号 + 文本 `1. ` | `1. 项` |
| `![说明](D:\本地\图.png)` | 真图片（OE + `one:Image`）；`![560](图.png)` 用 alt 位写宽度（pt） | `![图片]` 占位（OneNote 不回传文件路径） |
| `\| 表格 \|` | 真 OneNote 表格 | Markdown 表格 |

**正文基准字号**按这个顺序解析：目标页面自己的正文样式（`QuickStyleDef name="p"`，其次取出现最多的字号）→ 同分区取样（≤5 页）→ 注册表 `DefaultFontSize`（只读）→ 回落 12pt。`onenote_create` 的返回里带 `baseFontSize`，便于确认本次基准。

读回时按「相对正文的字号差」识别标题：优先匹配 `正文+headingDelta`（默认 +2/+1）与旧版
`正文+4/+2` 两套候选字号，再回退到 `≥正文+3` 认 `##`、`≥正文+1` 认 `###`；旧版本写的
`━━━ 标题 ━━━`、`▌ 标题` 仍然兼容读成 `##`/`###`。代码块在标题判定之前按
Consolas/monospace 字体识别（正文尺寸），所以标题下方的代码行不会被误判成 `##`。

> 手工在 OneNote 里改过字体/样式的块不会被插件"归一化"：插件只按配置写**新生成**的内容，
> `onenote_update` 会整页重写、自然归一；`onenote_append` 只追加，既有块保持原样。
> 需要统一时用 `onenote_read` + `onenote_update` 重刷一遍。

`onenote_read(withIds=true)` 的模型可见输出会在末尾附 `blocks (objectId):` 列表（CC 与 DSH 均生效），可直接把其中的 `objectId` 交给 `onenote_edit`。

## 已知限制

1. 编号列表只能用「圆点 + 序号」（`one:Number` 报 `0x80042001`，`bullet="0"` 报 `0x80042000`）。
2. `onenote_update` 会删掉页面内所有文本 Outline；其中的图片/墨迹/附件由插件**自动搬回**（`preserveMedia`，默认开），但**位置是近似的**：按原锚点段落就近放，锚点文本被改写就落到页尾。需要像素级精确位置时用 `onenote_edit` 原地改，不要整页重写。
3. 代码块语言标记（```js 的 `js`）无法保存。
4. 首次调用慢：OneNote 没跑时 COM 会拉起它并同步，可能 1–3 分钟，超时重试即可。
5. COM 需要非沙箱：OneNote 的 COM 服务器在文件沙箱里以 `0x80080005 (CO_E_SERVER_EXEC_FAILURE)` 失败。
6. `onenote_publish` 目标路径要用反斜杠（正斜杠报 `0x80042006`）；扩展名必须与格式匹配（`html`→`.htm`，`mht`→`.mht`，`docx`→`.docx`，`pdf`→`.pdf`），插件会自动归一化；不匹配时报 `INVALID_TARGET_PATH`（OneNote `0x8004201A`）。PDF 导出永远白底。
7. 长 ASCII 图会被页宽折行（默认 A4 竖排正文宽 468pt ≈ 61 半角字符/行）。
8. **独立分区既不能改名也不能删除**：`UpdateHierarchy` 和 `DeleteHierarchy` 对独立 `.one` 分区都返回 `0x8004200B SECTION_READ_ONLY`；笔记本内的分区可改名。要移除独立分区只能删 `.one` 文件后重启 OneNote。
9. **新建分区会自带一个空白页**：`onenote_create_section` 默认删除它（仅当分区只有这一页且无正文时），`keepBlankPage=true` 可保留。
10. **`deleteBlock` 会连带删除嵌套子块**：删除一个标题（如 h3）时，其下缩进的 bullet/段落作为子块一并消失；需要保留就先删除或移走子块，或用 objectId 精确删除目标块。
11. **独立分区要出现在 `onenote_list` 里，需要登记到 `extraSections`**：`onenote_create_section` 打开的分区会在当前进程内自动登记，立即可见；跨重启仍需写进 `ONENOTE_EXTRASECTIONS` 或 profile 的 `extraSections`。
12. **回收站页读不了、也恢复不了**：`DeleteHierarchy(permanent=false)` 之后，该 pageId 在 COM 里就已经不可见（`GetPageContent` → `0x80042014`，`onenote_read` 返回 `PAGE_DELETED`），`GetHierarchy` 也列不出；本机实测笔记本与独立分区都一样，而且 COM 没有 undelete。**唯一恢复途径是插件自己的备份**：`onenote_delete` / `onenote_update` / `onenote_insert_image` 会在动页面之前写一份整页 XML（含图片 base64）到 `backupDir`，返回里的 `backupPath` 就是它。
13. **`lastModifiedTime` 对 API 写入不生效，只随"页面被 OneNote 重新加载"刷新**：`UpdatePageContent` 改完内容后，页面属性里的 `lastModifiedTime` 仍是旧值（实测：创建后连续两次写入 + 各等 20s，内容已存盘、`.one` mtime 在变，属性不变）；但如果页面被导航/切到前台（UI 或 `NavigateTo`），它会立即刷新（实测 08:33:52 → 08:33:57），`dateTime` 则始终是创建时间。所以它既不能证明"没改"，也不能证明"改了"——判断是否真的改了页面请看 `onenote_read` 返回的 `revision`（页面 XML 的 SHA-1 前 16 位），`onenote_update` 也会返回写完后的 `revision`。
14. **`UpdatePageContent` 的删除语义**（实测，容易搞反）：patch 里**整段省略某个 Outline** → 该 Outline **不删**；patch 里带上某个 Outline（`objectID` 相同）但**省略其中的 OE** → 被省略的 OE **会删**。所以「整页重写」用 `DeletePageContent` + 新建，`onenote_edit` 用「重发所属 Outline 的完整内容」。

## 导入 markdown 文件（按章节拆页）

`dev/import-md.mjs` 把一个 md 文件导入为**独立分区**，每个 `##` 大章节一页：

```powershell
node dev/import-md.mjs <源 md 路径> <目标 .one 路径>
```

- 页面尺寸跟随 `ONENOTE_PAGESIZE`，默认不设（OneNote Automatic）；需要宽版时设成
  `{"width":842,"height":595,"landscape":true,"margins":{"top":36,"bottom":36,"left":36,"right":36}}`
- 幂等：已存在的同名页跳过；建分区时自动生成的空「未命名」页会被删掉
- 写入前剥离 `**粗体**`、`` `行内代码` ``、`[文本](url)` 链接（OneNote 子集不支持）
- 目录里的锚点链接会变成纯文本

## 测试

```powershell
# 纯函数单测 + MCP 协议测试（fake bridge，不碰 OneNote，沙箱可跑）
npm test

# 真机冒烟：list → create → read → delete，需要非沙箱 + OneNote
npm run smoke

# DSH 真机契约（需要本机 DSH 安装目录）：
#   $env:DSH_ELECTRON="<...>\DeepSeek Harness.exe"
#   $env:DSH_ASAR_MODULES="<...>\resources\app.asar\dsh\node_modules"
node dev/host-contract.mjs
```

COM 逻辑全在 `scripts/onenote-bridge.ps1`（纯 ASCII，中文走 JSON）。

### 保留空白块（tree / diagram）

围栏语言为 `tree` / `diagram` / `graph` / `art` / `ascii` / `pre` 时，按"保留空白块"处理：

- **0.7.6 起默认按正文段落渲染**（微软雅黑、正文字号、`margin:0in`、行间紧贴），不再用 Consolas；要旧的等宽外观设 `treeStyle: 'mono'`；
- 缩进换算成 U+3000（OneNote 会吃掉行首普通空格），层级由块内缩进步长推导，`n 级 = n 个 U+3000`；
- 前导 U+3000 原样保留（幂等），`├ ─ │ └ ┬ ┌ ┐` 原样写入；
- 不做 `//` 注释对齐（`alignCommentColumn` 对该块无效）、不加 `codePrefix`；
- 读回时命中盒线字符 + U+3000 缩进的段落，输出为 ```` ```tree ```` 围栏；
- 普通 ```` ``` ```` 代码块行为不变。

示例：

````markdown
```tree
pci_host_probe()                    // probe.c:3143
  ├─ pci_scan_root_bus_bridge()     // probe.c:3239
  │    └─ pci_scan_child_bus()      // probe.c:3088
  └─ pci_bus_add_devices()
```
````

限制：**表格单元格**里的 `->` 会被 OneNote 转成全角 `-＞`（仅表格；正文与代码块不受影响）。


### 0.7.10：`│` 引导线并入同一缩进网格

0.7.9 只处理了"纯空白缩进"的行，带 `│` 的行仍沿用源里的空格数，导致**同层级不同 x**（实测 `│  └─ a1` 87.61pt vs 空白缩进的 `└─ c1` 92.90pt，差 5.29pt）。

0.7.10 把 `│` 也算成网格单元：一个引导格 = `│` + 填充（3×NBSP + U+2004 = 13.80pt），合计 22.30pt，比一个缩进单位 20.88pt 宽 1.42pt。同层级的两类行现在落在 1.44pt 以内：

```
├─ a     72.02        └─ c1   92.90
│   └─ a1 94.34       └─ c2  113.78   （每级 20.88pt）
│       └─ a2 115.22  ...
```

实现上把行首解析成"单元格序列"（缩进格 / 引导格），重排时逐格输出，因此幂等；源里的 `│  └─`、`│     └─` 这类手排宽度会被折算成格数（按块内 gcd 步长四舍五入）。

### 0.7.9：树形缩进按实测字宽选单位（每级 ≈ `├─ ` 宽）

0.7.6 用 1 个 U+3000/级（12pt）比 `├─ `（20.51pt）窄 8.5pt，层级越深越往左叠。0.7.9 把缩进单位改成可配置的字符序列，默认 `bar`：

- `bar`（默认）：`\u00A0`×5 + `\u2004` = **20.88pt/级**（PDF 实测），与 `├─ ` 的 20.51pt 只差 0.37pt/级，6 级累计 +2.2pt；
- `nbsp`：6×NBSP = 21.29pt/级；
- `ideographic`：1×U+3000 = 12pt/级（0.7.6 行为）。

真机 6 级嵌套实测（PDF 逐字符 x 坐标，页面左边距 72.02pt）：

```
└─ c1  92.90    └─ c4  155.54
└─ c2 113.78    └─ c5  176.45
└─ c3 134.66    → 每级 20.88pt，c1 距父级 ─ 结尾约 0.4pt
```

读回时按"单位重复次数"识别层级（幂等）；旧页面里 1 个 U+3000 = 1 级的写法仍能正确换算。

### 0.7.7：`├─` / `└─` 后面固定一个半角空格

树行里的连接符后面统一成 **恰好一个 U+0020**：

- 源里漏了空格（`└─func()`）→ 补成 `└─ func()`；
- 源里是 NBSP（`└─\u00A0func()`，旧生成器为了让 HTML 不折叠而塞的）→ 归一成普通空格；
- 多个连字符/空格（`├──  y`）→ 归一成 `├─ y`；
- 幂等：已经是 `└─ x` 的行不变化。

配合 0.7.6 的"纯段落"写法（树行不切 Consolas span），旧版那种"空格落在 Consolas run 开头被折叠成 `└─func()`"的现象不会在新写入的内容里出现。

### 0.7.6：树形结构改为正文段落（默认），并修 3 个 edit 缺陷

树形图（` ```tree ` 围栏或直接写 `├─`/`└─`）不再走代码块，默认按**正文段落**渲染，导出 HTML 是：

```html
<p style='margin:0in;font-family:微软雅黑;font-size:12.0pt'>　├─ a</p>
```

- 无 Consolas、无内联 span 拆分、`margin:0in`（行间紧贴），字号跟随正文基准；
- 缩进按**块内实际的空格步长**推导（取所有缩进宽度的最大公约数），输出严格为 `n 级 = n 个 U+3000`，线性不再漂移（3/6/9 空格 → 1/2/3 个 U+3000）；
- 已经是 U+3000 的缩进原样保留（幂等）；
- **列对齐是近似的**：微软雅黑是比例字体（12pt 下 `├`=8.50pt、`─`=8.53pt、空格=3.48pt、`│`=8.50pt、U+3000=12.00pt），`├─ ` 占 20.51pt。0.7.9 起用 `treeIndent: 'bar'`（5×NBSP+U+2004 = 20.88pt/级）逼近，误差 0.37pt/级；要绝对精确只能整行等宽（`treeStyle: 'mono'`）；
- 想恢复旧外观：`treeStyle: 'mono'`（全局配置，或 `ONENOTE_TREE_STYLE=mono`）；
- 读回时按 U+3000 缩进 + 盒线字符识别为 ```tree 围栏；根行（无盒线）紧挨树行、同一 Outline 且不以句末标点结尾时并入围栏（0.7.8）。

同时修了 `onenote_edit` 的三个缺陷：

1. **`replaceText` 丢格式**：OneNote 把 run 的字号/字体放在 `<one:T>` 里的内联 `<span style=...>` 中，原实现直接替换 CDATA 导致 13pt 标题掉回 12pt。现在按替换前的实际字号/字体重新包 span。
2. **`replaceBlock` 忽略 `kind` 且生成非法 XML**：以前把新 `<one:T>` 追加到 `<one:OEChildren>` 之后（schema 要求 T 在子块之前），且 `kind` 完全没用。现在按 `kind` 重建（h2/h3 用 headingDelta + 段距，bullet/number 重建 `one:List`，para 只写正文字号），新 T 插在子块之前，子块保留。
3. **`insertAfter` 标题字号硬编码 4/2**：和 `headingDelta`（默认 2/1）不一致，插入的 h2 会变成 16pt。现在统一用 `headingDelta` + 对应段距。

### 0.7.5：存量表格字号修复工具

0.7.4 只保证**新写入**的表格字号正确；在此之前写入的页面，单元格仍停在 OneNote 表格默认的 13pt。0.7.5 新增 `onenote_normalize_tables`：

- 扫描页面里所有表格的单元格段落，把字号钉到该页正文基准（同时改 OE 的 `style` 和单元格文本里的内联 `span` 样式，OneNote 有时把格式放在 span 上）；
- 只重发包含表格的 Outline，其他内容逐字保留；改前默认备份整页 XML（返回 `backupPath`）；
- `dryRun=true` 只统计 `tables` / `cellsScanned` / `cellsChanged` / `outlinesUpdated`，不写页面；
- 返回 `baseFontSize`、各计数、`backupPath` 与 `warnings`。

实测：14 个存量页面的表格 `13.0pt` 计数全部归零，HTML 导出中 `<td>` 与正文 `<p>` 字号一致。

### 0.7.4：表格字号与正文一致

`one:Table` 的单元格 OE 之前没写字号，OneNote 会按**表格默认字号**（实测 13.0pt）落盘，于是正文 12pt 的页面上表格"大一号"（10.5pt 的页面上更是明显）。现在每个单元格段落都显式写 `font-size:<正文基准>pt`，和正文/列表用同一基准。

实测（同一页面）：正文 10.5pt 时表格单元格 10.5pt；正文 12pt 时单元格 12pt。

> 注意：这条只影响**新写入**的表格。已有页面里的表格仍是旧字号，需要重新生成该页（`onenote_update` 用原文 markdown）才会按新基准落盘。

### 0.7.3：同一锚点连续插图按调用顺序排列

`onenote_insert_image` 用文本锚点（`find`）时，如果锚点后面已经有图，新图会**排到那组图的末尾**；用 `objectId` 锚点则保持严格语义（紧贴该块插入）。这样同一个段落下连续插 3 张图，顺序就是调用顺序，不用倒着插。

- 文本锚点：`find:"锚点段落"` 连续插 100/200/300 → 页面顺序 100,200,300
- 块锚点：`objectId:<尾部段落>` + `position:"after"` → 仍然紧贴该段落

补充：锚点块自己带子块时（例如 h2 下面挂着几条 bullet），`after` 插在该**子块整体之后**，不是插进子块中间——这是层级语义，不是 bug。

### 0.7.2：回归里发现的 4 个问题

来自 lead 的 0.7.0 回归报告，逐条修好并真机复验（10/10）：

1. **`preserveMedia:false` 变成静默丢图**：之前只在"要保留"时才采集 media，关掉后 `preservedMedia`/`warnings` 全空。现在**总是先采集**，关掉时返回逐条 `placed:"dropped"` + `warnings: ["preserveMedia=false: N media item(s) were removed by this update"]`。
2. **缺图占位符继承了标题字号**：`[missing image] …` 的 OE 没写字号，嵌在标题下就继承 14pt，读回被当成 `## 标题`。现在显式写正文字号。
3. **同一锚点的多张图回插顺序反了**：每条 media 都重新找"第一个匹配锚点"再 `InsertAfter`，结果后一张插到前一张前面。现在记住每个锚点上最后插入的节点，后续同锚点依次排在它后面（原文档顺序）。
4. **`onenote_insert_image` 的返回值**：`width`/`height` 改成数字（原先是 `"600"` 这种字符串）；`objectId` 改成**写后回读到的真实 ID**（OneNote 会给新 OE 重新分配 ID，原来返回的是插件临时生成的、拿去做锚点会失败）。

文档补充：图片块的 `objectId` 在**整页 `onenote_update` 之后会变**（整页重建），插入/编辑后要重新 `onenote_read` 取当前 ID；同一 Outline 的部分编辑（`onenote_edit`）不会重编号。

### 0.7.1：read 的渲染输出带上 revision / media

`onenote_read` 的结构化返回一直有 `images` / `media` / `revision`，但**模型可见的是 `renderResult` 的文本**，之前页脚只有 `pageId` + `lastModified`，所以 agent 看不到"这页有几张图""revision 是多少"。0.7.1 把页脚改成：

```text
---
pageId: {...}
revision: ab63dce549e1feba
media: images=2 ink=0 files=0
lastModified: 2026-11-16T08:36:10.000Z
```

MCP 宿主（`server.js`）的文本结果同步改了。要拿每张图的 `objectId`，仍然用 `withIds=true`（图片块也会列出来）。

### 0.7.0：图片不再被静默删除，能插图了

这一版处理 5 个实测问题（来自 Harness lead 的现场报告）：

1. **`onenote_update` 不再静默删图**：重写前把整页 XML（含图片 base64）备份到 `backupDir`，返回 `backupPath`；重写后把原来的图片/墨迹/附件从备份里取出来重新插回，优先放回原锚点段落旁边，锚点文本没了放末尾。返回 `preservedMedia` 列表（类型/宽高/锚点/落点）和 `mediaCount`；万一没插全，`warnings` 会明说，备份里仍是完整的。要明确丢图就传 `preserveMedia: false`。
2. **新增 `onenote_insert_image`**：往已有页面插一张本地图片，可指定锚点块（`objectId` 或唯一 `find` 文本）、`after`/`before`/`end`、宽度（pt，等比缩放）。只重写锚点所在 Outline，页面其余部分不动；插入前同样备份。
3. **markdown 支持图片**：`onenote_create` / `onenote_update` / `onenote_append` 都能写独立一行的 `![说明](D:\本地\图.png)`；`![560](图.png)` 用 alt 位指定宽度，`ONENOTE_IMAGE_MAX_WIDTH`（默认 600pt）管默认宽度。文件不存在时写入 `[missing image] 路径` 占位并在 `warnings` 里报，不会让整页失败。
4. **`onenote_read` 报告图片**：`images` 数组（类型/宽/高/objectId/outlineId）+ `media` 计数 + markdown 里的 `![图片]` 占位；配合 `revision` 可以在 update 前后做丢失检测。
5. **`lastModified` 不再误导**：实测 OneNote 的 `lastModifiedTime` 对 `UpdatePageContent` 不刷新（始终是创建时间）。新增 `revision`（页面 XML 的 SHA-1 前 16 位），才是可靠的"页面变没变"判据。
6. **删除/回收站**：`onenote_delete` 删之前先备份整页 XML（含图片），返回 `backupPath`；删除后再读该页会得到明确的 `PAGE_DELETED`（而不是含糊的 `OBJECT_NOT_FOUND`）。实测 OneNote 回收站里的页 **COM 既读不到也恢复不了**（`0x80042014`，`GetHierarchy` 也不列），所以插件自己的备份就是唯一退路。

顺带实测澄清一条容易被搞反的语义（写进限制 14）：`UpdatePageContent` 里**整段省略一个 Outline 不会删除它**，但**带上该 Outline 却省略其中某个 OE，那个 OE 会删**。

### 0.6.8：段落字号与代码识别修复

- 正文/列表段落显式写入 base 字号：嵌套在标题下时不再继承标题字号，读回不会被误判成标题；
- 代码识别会检查 CDATA 内的内联 `span` 样式：OneNote 对中英混排的代码行会把字体移到 span 上（OE 本身没有 style），之前该行会被踢成正文、把一个代码块拆成几段；
- 提醒：`alignCommentColumn: 'auto'` 仍会重排**普通代码块**里的 `//` 注释列；`tree`/`diagram` 保留块不受影响。需要原样保留普通代码的注释列时，把 `alignCommentColumn` 设为 `0`。

