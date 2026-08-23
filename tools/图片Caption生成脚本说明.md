# 图片 Caption 生成脚本说明

## 1. 功能概述

`image_caption_generator.py`读取 Excel 第一个工作表中的图片，调用阿里云百炼 `qwen3.7-plus`识别画面并生成 Caption。

模型返回值由 API 的 `response_format=json_schema`、`strict=true`和`additionalProperties=false`约束，不是只在提示词中要求模型输出 JSON。

脚本不会覆盖源 Excel。结果、普通日志和完整结构化 JSONL 日志会保存到源文件旁的`caption_output`目录。

## 2. Excel 格式

第一行为标题，从第二行开始读取数据。只处理第一个工作表。

| 列 | 内容 | 说明 |
|---|---|---|
| A | 图片唯一标识 | 必填且不能重复；建议在 Excel 中存为文本 |
| B | URL或Data URI | 支持HTTP、HTTPS、`data:image/...;base64,...` |
| C | 本地图片绝对路径 | B列失败后才读取C列 |
| D | Caption描述 | 模型生成的最终Caption写入此列 |

图片读取顺序固定为：

```text
B列 URL/Data URI → C列本地绝对路径 → 失败交互
```

B列成功后不会继续读取C列。无论原始来源是什么，脚本都会校验图片并统一转换成`data:image/...;base64,...`后再发送给模型。

注意：Excel单个单元格最多保存32,767个字符，因此较大的图片不适合直接以Data URI放进B列。大图片请使用URL或C列本地路径。

## 3. 模型输出结构

每张图片的模型结果固定为：

```json
{
  "image_id": "图片唯一ID",
  "image_info": "主体、动作、场景等客观语义信息",
  "caption": "最终Caption描述"
}
```

约束包括：

- 只能出现以上三个字段；
- 三个字段全部必填且均为字符串；
- `image_id`使用单元素`enum`约束，只能返回当前A列ID；
- `image_info`和`caption`不能为空；
- 只有`caption`写入Excel的D列；完整对象写入JSONL日志。

如果百炼服务端违反根节点`type=object`约束，返回形如`[{...}]`的单元素对象数组，脚本会在`.log`中记录`MODEL_SCHEMA_SERVICE_VIOLATION`，去掉数组外壳后再次执行全部字段、类型、ID和非空校验。只有内部对象完整符合原JSON Schema时才会继续处理；多元素数组、空数组或其他类型仍会失败。该机制只是服务端异常的本地兜底，不会替代请求中的`response_format=json_schema`严格约束。

### 数组返回处理规则

每次模型请求只包含一张图片，因此最终只能对应一个结果对象。处理规则如下：

- 单元素对象数组`[{...}]`：移除唯一一层数组外壳，再执行完整Schema校验；校验通过后可写入D列。
- 空数组`[]`：判定失败。
- 多元素数组`[{...}, {...}]`：判定失败，不会默认采用第一个或最后一个元素。
- 单元素非对象数组（例如`["caption"]`）：判定失败。

多元素数组的错误示例：

```text
模型结果应为对象，服务端却返回数组：数组长度=2；仅单元素对象数组可安全归一化
```

发生上述失败后，脚本会把完整模型输入、原始响应和校验错误写入`.log`，并在自动重试提示中要求模型修正。自动重试最多共调用三次；仍然失败时，脚本向用户提供`[R]重试`、`[K]重新输入Key`、`[S]跳过`和`[Q]保存退出`。校验失败的数据行不会写入Excel的D列。

## 4. Caption规范

读取Excel成功后，脚本询问Caption描述规范。直接按回车使用默认值：

```text
按照【主体描述】+【修饰词】+【细节补充】+【风格/艺术形式】的逻辑，
整合成一段自然、连贯、客观的中文描述；不要输出栏目标签；控制在80至121字。
```

默认规范会在本地再次检查Caption长度。计数时忽略空白字符，中文标点计入；不在80至121字范围内时最多自动重试三次。

输入自定义规范后，脚本使用自定义规范，不再自动执行默认的80至121字长度检查。

## 5. 安装环境

推荐使用Python 3.10或更高版本。

在脚本目录打开PowerShell，执行：

```powershell
python -m pip install -r requirements_caption.txt
```

依赖如下：

- `openai`：通过OpenAI兼容接口调用阿里云百炼；
- `openpyxl`：读取和保存Excel；
- `Pillow`：检查图片格式、尺寸和有效性。

## 6. 先执行图片检查

建议正式调用模型前先运行检查模式：

```powershell
python image_caption_generator.py --check-only "D:\示例项目\input\图片数据.xlsx"
```

检查模式会：

- 读取第一个工作表；
- 汇总数据行、图片来源、空ID和重复ID；
- 按B列、C列顺序验证每张图片；
- 显示图片格式、尺寸和字节数；
- 不询问API Key；
- 不调用模型；
- 不修改Excel。

若URL存在防盗链、登录要求、Cloudflare验证或过期签名，检查时可能返回401或403。此类URL应更换为真正可公网直接下载的地址，或使用C列本地路径。

## 7. 正式运行

交互询问Excel路径：

```powershell
python image_caption_generator.py
```

也可以直接传入路径：

```powershell
python image_caption_generator.py "C:\完整路径\图片数据.xlsx"
```

正式运行流程：

1. 读取Excel并显示汇总；
2. 用户确认读取结果；
3. 用户选择默认Caption规范或输入自定义规范；
4. 隐藏输入阿里云百炼API Key；
5. 通过北京公共地址逐行调用`qwen3.7-plus`；
6. 严格解析JSON Schema结果；
7. 把`caption`写入D列；
8. 每成功一行便保存一次进度。

API Key只保存在当前进程内存中，不会写入代码、Excel、`.log`或`.jsonl`文件。

脚本固定使用北京公共地址，不询问业务空间ID：

```text
https://dashscope.aliyuncs.com/compatible-mode/v1
```

## 8. D列已有内容

每遇到一个D列非空的单元格，脚本都会询问：

```text
[O] 覆盖
[S] 跳过
[Q] 保存当前进度并退出
```

样例文件D列中的默认描述规范也属于“已有内容”，因此正式运行样例时会逐行询问是否覆盖。

## 9. 失败处理

### 图片获取失败

B、C列均失败时可选择：

```text
[R] 修改并保存源Excel后重新检查当前行
[S] 跳过
[Q] 保存当前进度并退出
```

选择重新检查后，脚本会重新打开源Excel，只读取当前行B、C列，不会丢失已经生成的Caption。

### 模型调用失败

接口错误或结果验证失败时可选择：

```text
[R] 重试
[K] 重新输入API Key
[S] 跳过
[Q] 保存当前进度并退出
```

网络错误、429和服务端5xx错误会先进行有限次数的自动重试。脚本没有设置`max_tokens`，避免严格JSON在输出中途被截断。

## 10. 输出文件

输出目录：

```text
源Excel所在目录\caption_output\
```

其中包括：

```text
原文件名_caption结果_时间戳.xlsx
原文件名_caption_时间戳.log
原文件名_caption_时间戳.jsonl
```

`.log`会逐次记录完整的模型输入、原始模型响应、接口错误响应和本地校验错误。模型输入包括完整图片Data URI，因此日志可能较大，也等同于包含图片原始内容，请妥善保管。API Key位于请求头，不会写入日志。

`.jsonl`用于汇总每行的图片ID、来源、尺寸、API响应ID、耗时、Token用量和最终结构化结果；其中仍会隐藏Data URI原文，并移除URL的查询参数和片段。

每一条Excel数据都会创建一次独立的无状态模型请求，只发送当前行的`system`消息、用户规范、图片ID和图片，不携带其他数据行的对话历史。同一行触发自动重试时会再次独立请求，并在当前提示中加入该行的纠错说明。

## 11. 当前限制

- 仅支持`.xlsx`和`.xlsm`；
- 仅处理第一个工作表；
- 仅处理A至D列；
- 支持JPEG、PNG、WEBP、GIF、BMP和TIFF；
- 单张图片最大20MB；
- C列必须为绝对路径；
- 不支持需要登录、Cookie、防盗链或浏览器验证才能下载的URL；
- 不支持加密或设置打开密码的Excel；
- 图片ID若包含前导零，应在Excel中存为文本，例如`00123`，不要存为数值格式。

## 12. 官方参考

- [qwen3.7-plus模型信息](https://help.aliyun.com/zh/model-studio/qwen3-7-plus)
- [阿里云百炼结构化输出](https://help.aliyun.com/zh/model-studio/qwen-structured-output)
- [阿里云百炼OpenAI兼容接口](https://help.aliyun.com/zh/model-studio/compatibility-of-openai-with-dashscope)
