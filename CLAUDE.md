# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Common Commands

1. **Initialize environment**: `pnpm install`
2. **Typecheck**: `pnpm typecheck`
3. **Lint**: `pnpm lint`

## DingdocsScript Instructions

- @knowledges/核心知识.md 是`DingdocsScript` 的 API 结构说明
- @knowledges/Base模块.md 是 `DingdocsScript` API 的主入口，提供了操作AI表格基础对象的方法，支持查询权限、管理数据表以及获取选择区域等信息。
- @knowledges/Sheet模块.md 负责维护数据与数据之间的联系，接口提供了操作数据表的方法，包括基本信息管理、字段管理、视图管理以及记录操作等功能。通过 Base 获取到 Sheet 之后，就可以调用 Sheet 中的 API，可以通过 base.getActiveSheet 方法来获取当前选中的数据表实例。
- @knowledges/Field模块.md 提供了操作数据表字段的方法，包括获取字段基本信息、修改字段属性以及操作字段值等功能。
- @knowledges/文本字段.md 专门用于处理文本字段数据。
- @knowledges/数字字段.md 专门用于处理数字字段数据。支持整数、小数，并提供多种数字格式化选项，如百分比、货币等。
- @knowledges/单选字段.md 专门用于处理单选类型数据。用于从预设选项中选择单个值，常用于状态、类别等场景。
- @knowledges/多选字段.md 专门用于处理多选字段类型数据，用于从预设选项中选择多个值。
- @knowledges/链接字段.md 用于存储和管理URL链接，可以包含显示文本和实际链接地址。它适合用于存储网站链接、资源地址等需要跳转的信息。
- @knowledges/日期字段.md 专门用于处理日期和时间类型数据。以毫秒级Unix时间戳形式存储。它支持不同的日期格式化选项，适合用于存储截止日期、创建时间、开始/结束时间等场景。
- @knowledges/货币字段.md 专门处理货币值，支持不同货币类型（如人民币、美元、欧元等）和不同精度的小数位数。它提供了专门的货币格式化和显示功能，使货币数据在表格中更直观。
- @knowledges/附件字段.md 用于存储和管理各种文件附件，如图片、文档、音频、视频等。它提供了上传、获取和管理附件的功能，支持多种文件格式。
- @knowledges/主键文档字段.md 专门用于处理主键文档字段数据。
- @knowledges/单向关联字段.md 专门用于处理单向关联字段数据。
- @knowledges/双向关联字段.md 专门用于处理双向关联字段数据。
- @knowledges/Record模块.md 表示数据表中的单条记录，提供了访问和操作记录数据的方法。

## Dingdocs.base Instructions

- @knowledges/Event 事件模块.md 处理事件相关的 API,可以在 UI 组件中通过`Dingdocs.base.event`直接使用。
- @knowledges/其他接口.md 获取地区、语言等国际化信息;获取当前主题。可以在 UI 组件中通过`Dingdocs.base.host`直接使用

## Code Architecture

项目有如下模块和功能，用于实现边栏插件，你应该确保代码逻辑都实现在`@src/script/service.ts` 和 `@src/components/App.tsx` 中:

1. **Model Service**: 位于 @src/script/service.ts, 包含插件的核心功能。相关功能利用 `DingdocsScript` API 实现相关业务逻辑. 并通过 `Dingdocs.script.registerScript` 注册相关方法，暴露给 UI 层调用。此代码运行在 web worker 中，不能使用各种浏览器 DOM/BOM 方法。

2. **UI Components**: 位于 @src/components/App.tsx, 用于渲染插件 UI. 组件可以 `Dingdocs.script.run` 方法调用 Model Service 中注册的方法。也可以使用 `Dingdocs.base.event` 注册各种事件。

3. **Locale**: 位于 @src/components/locales.ts, 包含 UI 上使用的所有国际化文案。

4. **注意**: service.ts 和 ui.tsx 完全是独立的两个文件，视图组件运行在 ui iframe 中，而 service.ts 运行在 script-iframe 中。两个iframe只能通过`Dingdocs.script.run`关联，由ui层调用service层注册的方法。


## Workflow

- 请先阅读本文档了解，如何进行项目开发。
- 所有 UI 上的文案需要支持国际化。
  - 支持语言包括 "zh-CN","ja-JP","en-US"。
  - 可以在ui.tsx里通过`Dingdocs.base.host.getLocale()` 获取当前语言
- 在异步获取的关键数据返回前，请展示 loading 页面。避免 UI 抖动。
- 完成任务后请使用 typecheck 和 lint 进行代码正确性检查。
