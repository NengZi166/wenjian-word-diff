# 文鉴 Wenjian

一个完全在浏览器本地运行的 Word `.docx` 文档比对工具，面向内网和隐私敏感场景。

选择原始版本和修改版本后，文鉴会：

- 生成一份带 Word 原生修订痕迹的 `.docx`；
- 在网页中预览文字、表格等差异；
- 提取并排版 Office 原生公式，列出公式新增、删除和修改；
- 检查 MathType/OLE、图片等嵌入对象是否变化；
- 全程在浏览器内存中处理文件，不上传、不登录、无遥测。

## 公式支持

| 公式类型 | 支持程度 | 说明 |
| --- | --- | --- |
| Office 原生公式（OMML） | 结构级比较 | 支持常见分式、上下标、根式、积分、求和、矩阵、定界符等结构，并用 KaTeX 在网页排版 |
| MathType / OLE | 对象级比较 | 可判断对象新增、删除或替换；暂不解析对象内部单个符号 |
| 图片公式 | 图片级比较 | 可判断图片文件是否变化；不进行 OCR |

Word 修订文件由 Docxodus 的 `WmlComparer` 引擎生成。项目上游包含多组原生公式、图片与公式混排、多个公式等测试样例。

## Windows 离线使用

从 Releases 下载离线压缩包，完整解压后双击：

```text
启动文鉴.exe
```

浏览器会自动打开本地地址。黑色启动窗口需要保持打开；关闭该窗口即停止工具。启动器不调用 PowerShell、CMD、Node.js 或 Python，离线包也不要求安装 Word 或 WPS。

## 从源码运行

需要 Node.js 22.13 或更高版本：

```bash
npm install
npm run dev
```

浏览器访问终端显示的本地地址。

生成 Windows 离线目录：

```bash
npm run package:offline
```

成品位于 `offline-dist/`。它是纯静态站点，可放到内网 IIS、Nginx 或其他静态服务器；服务器需要为 `.wasm` 返回 `application/wasm`。

## 隐私模型

- 文件只通过浏览器 `File` API 读取；
- DOCX 解包、公式分析和 Word 比对都在浏览器内完成；
- 应用没有上传接口、账号系统或分析埋点；
- 刷新或关闭页面会清除已选文件和生成结果；
- 如部署到内网服务器，服务器只提供静态程序文件，不接收用户文档。

## 当前限制

- 仅支持 `.docx`，不支持旧版 `.doc`；
- 不支持加密或设有打开密码的文档；
- 单个文件上限为 100 MB；
- 网页预览不等同于 Word 的排版引擎，最终结果以下载的修订版 Word 为准；
- MathType/OLE 和图片公式暂时无法定位到内部单个符号。

## 技术栈

- React + TypeScript
- Docxodus WebAssembly：生成 Word 原生修订文件
- JSZip：读取 DOCX/OOXML 包
- KaTeX：网页公式排版

## 开源协议

项目代码使用 [MIT License](LICENSE)。第三方组件的声明见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
