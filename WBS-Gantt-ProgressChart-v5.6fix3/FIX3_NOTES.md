# 5.6fix3 修复说明 / Repair notes

交付文件：`WBS-Gantt-ProgressChart-v5.6fix3.xlsm`；工程版本 `5.6.3`。

## 按钮随单元格移动

fix2 的 Milestones 按钮使用 `oneCellAnchor`，VBA 复位过程又设置 `xlMove`，即“随单元格移动、大小不变”。修改列宽或移动单元格后，Excel 会先移动按钮；后续初始化或刷新再恢复原坐标，因此出现“先偏移、点击后又复位”。Sheet1（WBS）使用 `absoluteAnchor`，对应“不随单元格移动或调整大小”，不依赖复位宏。

fix3 的两张表统一调用同一个固定坐标按钮构建函数。Milestones 的 11 个按钮改用 `absoluteAnchor`，形状坐标与锚点坐标一致；保存文件本身已指定固定位置。生成、切页和全刷新均不再改写工具栏位置。

| 控件 | Left（磅） | Top（磅） | Width（磅） |
| --- | --- | --- | --- |
| Generate、Copy、Widen、Narrow | 25、208、335、472 | 93 | 171、114、125、125 |
| Undo、Redo、Delete Level 1–5 | 25、121、217、325、433、541、649 | 235 | 88、88、100、100、100、100、100 |

高度均为 25 磅；交付布局与第 5、11 行顶端间距仍为 3 磅，横向顺序及间距沿用 fix2。固定坐标相对于工作表，含义与 Sheet1 相同；用户调整顶部行高时，不再自动跟随行顶端或恢复行高。

## 清理及兼容性

- 删除 `moving_anchor`、`move_row` 分支及不再使用的 `math` 导入。
- 删除 `Sheet4.ResetToolbarLayout`、`mLayingOut`、重复的名称/坐标/尺寸数组及初始化、Refresh All 内的调用。
- 删除旧的“人为堆叠后调用复位”测试，改为检查固定锚点，以及在 Excel 场景里改变行列和移动单元格后立即核对位置。补充脚本改为通用名称 `excel_regressions.ps1`，保留有效的输入错误回归。
- 更新预览读取方式、打包清单与文档，清除“随单元格移动并自动复位”的过期操作说明。
- 所有用户按钮名称、宏入口及绑定继续兼容；移除的 `ResetToolbarLayout` 是内部复位过程，没有按钮绑定。
- 保留错误 28 防重入、迭代格式快照、两张表各自三步历史、五级删除、日期/来源校验、失败规范化回滚及手动计算判重。

与 fix2 的 XLSM 包对比，仅 VBA、Milestones 图形和 WBS 版本标题发生变化。其余数据、样式、验证、WBS 工具栏和样例图形保持一致。依赖版本不变，锁文件仅更新本工程版本。原 fix2 交付文件保留，修改前工程已备份。

## 验证

详见 `CHECK_RESULTS.md` 或 `CHECK_RESULTS_zh-CN.md`。24 项文件回归通过；同一固定锚点检查可检出 fix2 的问题并通过 fix3。VBA 语法、OOXML、嵌入源码和交付包一致性另行核对。

Windows Excel 脚本准备了事件关闭/开启、行高列宽增减、单元格剪切移动、缩放、切页、生成、全刷新和临时副本保存重开场景。本环境没有桌面 Excel，这些运行场景未执行；预览不是 Excel 截图。

## English

Milestones previously used “Move but don't size with cells” and relied on later VBA calls to repair its coordinates. All 11 buttons now share Sheet1's absolute worksheet positioning (“Don't move or size with cells”). The cell-anchor conversion, runtime repair routine, redundant coordinate arrays and their callers were removed. Initial geometry and all user-facing macro bindings are retained. The regression checks enforce fixed anchors and matching shape geometry; desktop Excel scenarios are prepared but were not executed in Linux.
