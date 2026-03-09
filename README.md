# 清结算网银账单Excel生成小工具

基于 Electron + SQLite + XLSX 的桌面端工具，支持在 Windows 10 和 Windows 11 上运行。

## 功能

- 导入 Excel / CSV 作为模版文件。
- 基于 `COMMON枚举.xlsx` 维护模版字段与 COMMON 字段映射。
- 导入 Excel / CSV 账单文件并按映射替换表头。
- 生成 `模版名-COMMON-执行日期.xlsx` 文件到按日期创建的目录中。
- 支持另存为导出生成文件。
- 将模版和映射关系持久化到 SQLite。
- 异常按日期写入日志文件。

## 运行

```bash
npm install
npm run init:enum
npm start
```

## 打包 Windows 安装包

```bash
npm run dist:win
```

打包后的安装程序默认位于：

```bash
dist/清结算网银账单Excel生成小工具-1.0.0-setup.exe
```

## GitHub 下载说明

- GitHub 网页上的 `Download ZIP` 下载的是源码，不包含已构建的 `exe`。
- 如果需要现成安装程序，请到仓库的 `Actions` 页面下载 `windows-installer` 构建产物。
- 如果是在本地 Windows 机器上自行生成 `exe`，执行 `npm install` 后运行 `npm run dist:win`。

## 数据和日志目录

- 生成文件目录：`文档/清结算网银账单Excel生成小工具/exports/执行日期`
- 日志目录：`文档/清结算网银账单Excel生成小工具/logs`
- SQLite 数据库：Electron `userData` 目录下的 `tool-data.sqlite`

## 注意

- 仓库根目录的 `COMMON枚举.xlsx` 当前为示例枚举，请替换成实际业务枚举后再正式使用。
- 如果重复导入同名模版，系统会保留模版名称并重置旧映射关系，需重新维护映射。
