# 订单评价助手（Tauri 桌面版）

基于 `Tauri + React + Rust` 的本地桌面工具，面向你的场景：
- 导入 Excel（指定订单号列）
- 用聚水潭 Cookie 查询订单商品
- 商品名称匹配本地图片目录
- 调用 AI 生成评价（提示词可自定义）
- 为每个商品随机拷贝图片并输出汇总结果

## 已实现功能

- 本地配置持久化（`settings.json`）
- 打开聚水潭登录页（系统浏览器）
- Cookie 可用性校验（通过 `__VIEWSTATE` 验证）
- Excel 首个 Sheet 的订单列读取
- 聚水潭订单批量查询商品（含缺失订单单独重试）
- 图片目录别名匹配（支持 `目录名A##别名B`）
- 评价文案去重重试（同商品目录尽量不重复）
- 输出结果：
  - 订单目录（图片 + txt）
  - `summary.xlsx`（状态汇总）

## 运行

```bash
pnpm install
pnpm tauri dev
```

## 打包

当前（2026-03-06）可在 macOS 上构建 macOS 包：

```bash
pnpm tauri build
```

Windows 包建议在 Windows 机器或 Windows CI Runner 上构建（最稳妥）：

```bash
pnpm install
pnpm tauri build
```

## 使用说明

1. 选择 Excel 文件。
2. 设置订单列名（默认 `订单号`）。
3. 点击“打开登录页”，在浏览器登录聚水潭后，把 Cookie 粘贴回工具中。
4. 点击“校验 Cookie”。
5. 选择图片根目录：每个商品一个文件夹，文件夹名可写别名（`主名##别名`）。
6. 配置 AI Base / Key / 模型 / 提示词模板（支持 `{product_name}` 占位符）。
7. 点击“开始生成”。

## 输出结构示例

```text
输出目录/
  任务名_时间戳/
    订单号A/
      1_商品目录名.txt
      1_商品目录名_1.png
      1_商品目录名_2.png
      ...
    订单号B/
      ...
    summary.xlsx
```

## 注意事项

- Cookie 存在有效期，失效后需重新登录获取。
- AI Key 和 Cookie 都保存在本地配置文件中，请注意本机权限与备份安全。
- 如果某商品未匹配到图片目录，会在 `summary.xlsx` 里标记“未匹配到图片目录”。
