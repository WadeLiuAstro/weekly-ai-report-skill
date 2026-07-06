---
name: weekly-ai-report
description: "AI 应用情报周报自动化生成。覆盖新闻采集、项目筛选、OOXML 格式化、docx 输出全流程。当用户要求生成本周/指定周期的 AI 周报、向现有周报添加新闻或项目、或提到'周报''AI应用情报''NewsUpdate'等关键词时使用。"
agent_created: true
---

# weekly-ai-report — AI 应用情报周报生成

## 触发条件

当用户提出以下请求时，加载本 Skill：
- "生成本周 AI 周报"
- "生成一份本周的 AI 周报"
- "把这个星期的 AI 新闻做一份周报"
- 向现有 `NewsUpdate_Fri_*.docx` 添加新闻或项目
- 涉及"周报""AI应用情报""NewsUpdate"关键词的 docx 生成任务

## 周报结构

文件命名：`NewsUpdate_Fri_YYYYMMDD.docx`（周五日期）

两大板块，顺序固定：

```
[新闻类] 板块标题
  ├── 日期1 (如 2026/7/3)
  │   ├── 新闻标题1 (加粗斜体)
  │   ├── 正文摘要
  │   └── 来源链接 (HYPERLINK FIELD CODE)
  ├── 日期2
  │   └── ...
  └── ...

[项目类] 板块标题
  └── 5列表格
      | 公司名 | 一句话 | 融资 | 团队背景 | source |
```

---

## 一、内容采集规范

### 1.1 新闻类

**时间范围**：上周五 → 本周五（7天，可按用户要求调整）

**数量**：每日 2-4 条，整周 14-21 条

**新闻源**：
- 英文：aitoolsrecap.com、TechCrunch、The Information
- 中文：aitntnews.com、neican.ai、iaipie.com、36氪、新浪财经、微信公众号（AI闹、白鲸出海）

**优先级排序**：
1. 重大融资/IPO（如快手可灵30亿美元、DeepSeek 500亿元）
2. 政策监管（如白宫阻止GPT-5.6、出口管制）
3. 模型/产品发布（如GPT-5.6、美团LongCat-2.0）
4. 行业趋势/竞争格局（如Claude流量暴增、企业转向DeepSeek）

**新闻摘要撰写要求**：
- 每条正文 100-200 字，覆盖 5W1H（Who/What/When/Where/Why/How）
- 多条相关新闻可合并为一条（如"阿里整合三大Agent产品线"）
- 提及的具体数据（融资金额、估值、参数规模）必须准确

### 1.2 项目类

**筛选原则（关键）**：
- **必须符合 Garena 定位**：娱乐、泛娱乐赛道
- 重点赛道：AI社交、AI陪伴、AI短剧、AI音乐、AI游戏、AI视频生成、AI 3D生成、AI虚拟人
- **排除**：纯企业服务、机器人硬件、工业AI（除非有娱乐/游戏应用场景）
- 每周 5-7 个项目

**数据来源**：微信公众号（AI闹、白鲸出海）的创业项目介绍文章，辅以 36氪、新浪财经

**提取字段**：
| 字段 | 要求 |
|------|------|
| 公司名 | 中英文全称，如"快手可灵AI (Kling)" |
| 一句话 | 赛道+核心功能，1-2句话 |
| 融资 | 时间+轮次+金额+投资方+估值（无融资写"未披露"） |
| 团队背景 | 创始人、核心团队、孵化背景 |
| source | 信息来源，如"36氪/新浪财经" |

**红色标记规则**：有明确融资信息（金额、轮次、投资方）的公司名标红 `EE0000`；融资"未披露"的不标红。

---

## 二、OOXML 格式规范（精确 XML）

> **核心原则**：以最新一期周报（如 `NewsUpdate_Fri_20260703.docx`）为格式模板，逐属性复制。

命名空间：
```python
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
```

### 2.1 板块标题（"新闻类"/"项目类"）

```xml
<w:p>
  <w:pPr>
    <w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/>
      <w:color w:val="1F4E79"/>
      <w:sz w:val="32"/>
      <w:lang w:eastAsia="zh-CN"/>
    </w:rPr>
    <w:t>新闻类</w:t>
  </w:r>
</w:p>
```

关键属性：`w:b` | `w:color val="1F4E79"` | `w:sz val="32"` | **无 rFonts**

### 2.2 日期段落（如 "2026/7/3"）

```xml
<w:p>
  <w:pPr>
    <w:spacing w:after="0"/>
    <w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:color w:val="000000"/>
      <w:sz w:val="28"/>
      <w:u w:val="single"/>
      <w:lang w:eastAsia="zh-CN"/>
    </w:rPr>
    <w:t>2026/7/3</w:t>
  </w:r>
</w:p>
```

关键属性：`w:color val="000000"` | `w:sz val="28"` | `w:u val="single"` | **无 rFonts**

### 2.3 新闻标题

```xml
<w:p>
  <w:pPr>
    <w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/>
      <w:i/>
      <w:color w:val="000000"/>
      <w:lang w:eastAsia="zh-CN"/>
    </w:rPr>
    <w:t>新闻标题文本</w:t>
  </w:r>
</w:p>
```

关键属性：`w:b` | `w:i` | `w:color val="000000"` | **无 sz、无 rFonts**

### 2.4 正文段落

```xml
<w:p>
  <w:pPr>
    <w:rPr>
      <w:rFonts w:eastAsia="宋体"/>
      <w:kern w:val="2"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="24"/>
      <w:lang w:eastAsia="zh-CN"/>
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:eastAsia="宋体"/>
      <w:kern w:val="2"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="24"/>
      <w:lang w:eastAsia="zh-CN"/>
    </w:rPr>
    <w:t>正文内容...</w:t>
  </w:r>
</w:p>
```

**关键**：`w:rPr` 在 `pPr` 和 `r` 中**都要设置**，属性包括：`rFonts eastAsia="宋体"` | `kern val="2"` | `sz val="22"` | `szCs val="24"` | `lang eastAsia="zh-CN"`

### 2.5 来源链接（HYPERLINK FIELD CODE）

**⚠️ 必须使用 FIELD CODE，不能用简单 `<w:hyperlink>` 关系。**

```xml
<w:p>
  <w:pPr>
    <w:widowControl w:val="0"/>
    <w:spacing w:after="160" w:line="278" w:lineRule="auto"/>
    <w:rPr>
      <w:rStyle w:val="136"/>
      <w:rFonts w:eastAsia="宋体"/>
      <w:kern w:val="2"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="24"/>
      <w:lang w:eastAsia="zh-CN"/>
    </w:rPr>
  </w:pPr>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://..." \h </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r>
    <w:rPr>
      <w:rStyle w:val="136"/>
      <w:rFonts w:eastAsia="宋体"/>
      <w:kern w:val="2"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="24"/>
      <w:lang w:eastAsia="zh-CN"/>
    </w:rPr>
    <w:t>https://...</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rStyle w:val="136"/>
      <w:rFonts w:eastAsia="宋体"/>
      <w:kern w:val="2"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="24"/>
      <w:lang w:eastAsia="zh-CN"/>
    </w:rPr>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
</w:p>
```

关键属性：`widowControl val="0"` | `spacing after="160" line="278" lineRule="auto"` | `rStyle val="136"` | 显示 run 中同样需要 `rFonts eastAsia="宋体"` + `kern val="2"` + `sz val="22"` + `szCs val="24"`

**FIELD CODE 结构**：`begin → instrText("HYPERLINK URL \\h") → separate → r(display text) → end`

### 2.6 表格

**表头行**：
```xml
<w:tr>
  <w:trPr><w:jc w:val="center"/></w:trPr>
  <w:tc>
    <w:tcPr>
      <w:tcW w:w="1746" w:type="dxa"/>
      <w:shd w:val="clear" w:color="auto" w:fill="D9EAF7"/>
    </w:tcPr>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:rPr>
          <w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>
          <w:sz w:val="18"/>
          <w:lang w:eastAsia="zh-CN"/>
        </w:rPr>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>
          <w:b/>
          <w:sz w:val="18"/>
          <w:lang w:eastAsia="zh-CN"/>
        </w:rPr>
        <w:t>公司名</w:t>
      </w:r>
    </w:p>
  </w:tc>
  <!-- 其余 4 列结构相同，tcW 分别为 1723/1723/1720/1723 -->
</w:tr>
```

**列宽**：1746 / 1723 / 1723 / 1720 / 1723 (dxa)

**表头格式**：`w:shd fill="D9EAF7"` | `w:b`（加粗） | `w:rFonts ascii/hAnsi/eastAsia="宋体"` | `w:sz val="18"` | `w:jc val="center"`

**数据行**（每行格式相同，仅"公司名"列颜色不同）：

```xml
<w:tc>
  <w:tcPr><w:tcW w:w="1746" w:type="dxa"/></w:tcPr>
  <w:p>
    <w:pPr>
      <w:widowControl w:val="0"/>
      <w:jc w:val="center"/>
      <w:rPr>
        <w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>
        <w:sz w:val="18"/>
        <w:lang w:eastAsia="zh-CN"/>
      </w:rPr>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>
        <!-- 有融资：加 color EE0000 -->
        <w:color w:val="EE0000"/>
        <!-- 无融资：不加 color -->
        <w:sz w:val="18"/>
        <w:lang w:eastAsia="zh-CN"/>
      </w:rPr>
      <w:t>公司名</w:t>
    </w:r>
  </w:p>
</w:tc>
```

**表格边框**（tblPrEx 中）：
```xml
<w:tblBorders>
  <w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:left w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:right w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:insideH w:val="single" w:color="auto" w:sz="4" w:space="0"/>
  <w:insideV w:val="single" w:color="auto" w:sz="4" w:space="0"/>
</w:tblBorders>
```

**单元格边距**：
```xml
<w:tblCellMar>
  <w:top w:w="0" w:type="dxa"/>
  <w:left w:w="108" w:type="dxa"/>
  <w:bottom w:w="0" w:type="dxa"/>
  <w:right w:w="108" w:type="dxa"/>
</w:tblCellMar>
```

### 2.7 页面设置（sectPr）

```xml
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"
           w:header="720" w:footer="720" w:gutter="0"/>
  <w:cols w:space="720" w:num="1"/>
  <w:docGrid w:linePitch="360" w:charSpace="0"/>
</w:sectPr>
```

---

## 三、技术实现规范

### 3.1 工作流

```
1. 读模板 → 2. 采新闻 → 3. 筛项目 → 4. 生成XML → 5. 打包docx → 6. 验证格式 → 7. 交付
```

### 3.2 Python OOXML 操作步骤

```python
import zipfile, copy, os, shutil
import xml.etree.ElementTree as ET
from datetime import date, timedelta

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'

# 1. 确定参考模板：上一期周报（如 NewsUpdate_Fri_20260703.docx）
# 2. 解压 docx 为临时目录
# 3. 解析 word/document.xml
# 4. 构造新段落/表格行（按第二章格式规范逐属性设置）
# 5. 插入到正确位置
# 6. 写回 document.xml，重新打包为 .docx
# 7. 格式验证
```

### 3.3 插入位置规则

- **新闻类**：在"新闻类"板块标题之后、"项目类"板块标题之前插入
- **项目类**：在表格末尾（最后一个 `w:tr` 之后、`</w:tbl>` 之前）追加新行
- 新日期组：在日期段落之后插入空行（`<w:p/>`），再插入新闻条目

### 3.4 关键注意事项

1. **链接必须用 FIELD CODE**（`fldChar`/`instrText`），不能用简单 `<w:hyperlink>` 关系
2. **正文必须包含** `kern`/`szCs`/`rFonts eastAsia="宋体"`，否则与参考文档格式不一致
3. **链接段落必须包含** `rStyle`/`widowControl`/`spacing`
4. **表格行操作要保留** `tblPrEx`（边框和单元格边距）
5. **paraId 属性**：可选，但建议为每个段落生成唯一 ID 以兼容 Word 修订功能
6. **样式 136**：即 Hyperlink 字符样式，不要修改或删除

---

## 四、格式验证清单

生成 docx 后，必须逐项检查：

| # | 检查项 | 预期值 |
|---|--------|--------|
| 1 | 板块标题颜色 | `1F4E79`, `sz=32`, 加粗 |
| 2 | 日期格式 | `sz=28`, 下划线 `single` |
| 3 | 新闻标题 | 加粗 + 斜体, `color=000000` |
| 4 | 正文字体 | `eastAsia="宋体"`, `kern=2`, `sz=22`, `szCs=24` |
| 5 | 链接格式 | FIELD CODE, `rStyle=136`, `widowControl=0`, `spacing line=278` |
| 6 | 表格表头 | `fill=D9EAF7`, 宋体, `sz=18`, 加粗 |
| 7 | 有融资标记 | 公司名 `color=EE0000` |
| 8 | 项目赛道 | 全部属于娱乐/泛娱乐 |
| 9 | 新闻数量 | 14-21 条 |
| 10 | 项目数量 | 5-7 个 |

---

## 五、交付规范

每次完成周报生成后：

1. **deliver_attachments** 交付 docx 文件
2. **open_result_view** 预览文件
3. **追加工作记忆**到 `周报/.workbuddy/memory/YYYY-MM-DD.md`：
   - 覆盖周期、新闻数量、项目数量、新闻亮点
   - 格式验证结果
   - 遇到的问题和解决方案

---

## 六、用户偏好参考

以下偏好来自用户（Wade Liu）的长期工作习惯：

- **语言**：中文交流，中英术语混排
- **信息源**：优先微信公众号（AI闹、白鲸出海）、中文 AI 资讯站
- **项目聚焦**：Garena 定位 = 娱乐 + 泛娱乐，非此赛道的项目不纳入
- **格式一致**：严格匹配参考文档的 XML 属性，不做"近似"
- **原文档修改**：不覆盖源文件，生成新文件（如 `NewsUpdate_Fri_YYYYMMDD.docx`）
- **简洁表述**：公司描述精炼，1-2 句话说清定位
- **数据准确**：融资金额、估值、参数规模必须核实
