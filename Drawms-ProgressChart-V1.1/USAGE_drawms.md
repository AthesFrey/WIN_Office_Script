# drawms 使用说明

用桌面版 Microsoft Excel 打开 `drawms.xlsm` 并启用宏。工作簿只有 `Milestones` 一张表。

表格只有两列：`Milestone name` 和 `Milestone date`。两列预设为文本格式，输入内容按原样保存和绘制；日期不会被解析、改写、排序或合并。某行任意一列有内容就会生成一个节点，两列都为空的行跳过，最多生成 60 个节点。

点击 `Generate Chart` 生成图形。间距过小时，名称和日期会按 fix3 的避让规则分别上下错层，必要时显示引导线。`Widen Spacing` 和 `Narrow Spacing` 每次调整 10%，范围为 50% 到 200%。

drawms 没有 WBS 联动、导入、复制、层级删除、自定义撤销/恢复或快捷键接管。Excel 自带的 Undo/Redo 保持默认行为。
