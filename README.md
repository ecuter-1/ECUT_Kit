# 📊 ECUT 成绩助手

> 东华理工大学教务系统成绩增强脚本 — 一键计算 GPA、加权均分、导出 Excel，完整支持 iPhone / iPad 触控。

[![Version](https://img.shields.io/badge/version-4.1.0-blue.svg)](https://github.com/yourusername/ecut-grade-enhancer/releases)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Tampermonkey%20%7C%20Violentmonkey-orange.svg)]()

---

## ✨ 功能特性

| 功能 | 说明 |
|------|------|
| 📈 GPA 实时计算 | 平均学分绩点、算术均分、加权均分、总学分、总学分绩点、挂科数 |
| 🔍 智能筛选 | 一键剔除选修课 / 挂科课程后重新计算 |
| 📅 全学期汇总 | 自动切换"全部学期"并加载全部成绩后计算 |
| 🔢 列排序 | 点击任意表头按课程名称、学分、成绩、绩点等排序 |
| 📝 课程详情 | 点击任意行弹窗展示该课程所有原始字段 |
| 📥 导出 Excel | 自动以"姓名\_学期\_教务系统成绩明细.xlsx"命名导出 |
| 📱 移动端适配 | iPhone / iPad 触控拖拽、响应式布局、最小化浮窗 |
| 💾 状态记忆 | 窗口位置与大小自动持久化，刷新后保留 |

---

## 🚀 安装方法

### 第一步：安装脚本管理器

请先在浏览器安装以下任一扩展：

- **Chrome / Edge**：[Tampermonkey](https://www.tampermonkey.net/) 或 [Violentmonkey](https://violentmonkey.github.io/)
- **Firefox**：[Tampermonkey](https://www.tampermonkey.net/) 或 [Violentmonkey](https://violentmonkey.github.io/)
- **Safari (iOS/macOS)**：[Userscripts](https://apps.apple.com/app/userscripts/id1463298887)

### 第二步：安装脚本

**方式 A（推荐）** — 从 GreasyFork 一键安装：

> 🔗 *发布后在此处添加 GreasyFork 链接*

**方式 B** — 手动安装：

1. 打开 Tampermonkey 控制台 → "新建脚本"
2. 将 [`src/ecut-grade-enhancer.user.js`](src/ecut-grade-enhancer.user.js) 的全部内容粘贴进去
3. 保存（`Ctrl+S`）

### 第三步：使用

1. 登录东华理工大学教务系统
2. 进入 **成绩查询** 页面
3. 脚本自动启动，悬浮窗出现在页面右上角
4. 可拖动悬浮窗到任意位置，点击"收起"最小化

---

## 📖 使用说明

### 基本操作

- **🚀 重新分析**：手动触发数据扫描与重新计算
- **📅 全部学期**：自动将学年/学期切换为"全部"并重新查询，再次点击恢复原来学期
- **📥 导出Excel**：将当前所有成绩原始数据导出为 `.xlsx` 文件

### 筛选器

| 选项 | 效果 |
|------|------|
| 剔除选修 | 计算 GPA 时排除含"选"或"公"字的课程性质 |
| 剔除挂科 | 计算 GPA 时排除绩点 < 1.0 的课程 |

### 成绩映射规则

对于非数字成绩，脚本按以下规则转换为数值：

| 原始成绩 | 映射分值 |
|---------|---------|
| 优秀 / 优 | 95 |
| 良好 / 良 | 85 |
| 中等 / 中 | 75 |
| 及格 | 65 |
| 合格 / 通过 | 80 |
| 不及格 / 不合格 / 不通过 | 50 |
| 缓考 / 缺考 / 作弊 | 0 |

---

## 📱 移动端使用

- **拖拽**：按住顶部标题栏拖动，最小化时可拖动整个浮窗
- **横屏**：自动适配高度布局
- **最小化**：点击"收起"按钮，浮窗折叠为紧凑模式
---

## ⚠️ 免责声明

本脚本仅供学习交流使用，通过读取页面上已渲染的成绩数据进行本地计算，**不会修改、上传或泄露任何个人数据**。

使用本脚本所产生的任何后果由使用者自行承担，与开发者无关。请遵守学校相关规定。
