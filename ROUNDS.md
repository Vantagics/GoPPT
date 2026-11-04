# 保真度轮次 r75–r85笔记

从网上取得 79 个复杂度各异的 PPTX（LibreOfficecore 回归集、apache/poi 实拍文稿、
POI 特性文稿），用 cmd/zz_render + cmd/zz_cmp 与 PowerPoint COM 导出
（cmd/zz_cmp/export_office.ps1）逐像素对比，两阶段：渲染保真与生成保真
（cmd/zz_roundtrip：读入→写出→Office 导出对比 Office 导出原稿）。
每一轮归因到具体机制、修复并带回归测试提交。

## 本轮修复（提交号）

| 轮 | 修复 | 效果（vs Office 导出） |
|----|------|--------|
| r75 | 9180bb8 截断 zip 抢救 + 59bf574 画布高度取整 | Divino_Revelado 12/12 页可读；16:10 级画布逐像素配对 |
| r76 | 6706d0b 母版/版式非占位符装饰形状 | alterman_security 16.4% → <8%；红条带像素级一致 |
| r77 | c95a8c3 每母版一套主题 + showMasterSp="0" | themes（7 母版 7 主题）背景/文字纠正 |
| r78 | dff001f 显式 b="0" 压过母版 bold | ShapeLineProperties 换行点与粗细一致 |
| r79 | 613a146 图表：自动标题/右侧图例档/自动饼图半径 | pie-chart 8.6% → 4.2% |
| r80 | f135b1f 母版 clrMap 重映射颜色名 | master-bg-color/onemaster-twolayouts 100%/50% → 0.00% |
| r81 | 30257bb SmartArt 框架部件往返保留 | fill-color-list 生成侧 47.6% → 7.2%；customGeo 49 部件回写 |
| r82 | 942c61a 版式 overrideClrMapping | themes slide09 97.8% → 28.8% |
| r83 | 2b40e12 背景 blipFill `<a:tile/>` 平铺 | page_transparent_bitmap 22.1% → 0.95% |
| r84 | 6384348 图片占位符继承几何 | customshape-bitmapfill-srcrect 26.2% → 0.29% |
| r85 | e5fbea1 custGeom 多子路径 fill 属性 | treck 标记的未填充内区正确留空（余属位置差另日志跟进） |

## 复现工具

- `cmd/zz_corpus`：语料普查（形状种类、unsupported、无图数据）
- `cmd/zz_render`：渲染整包页为 slideNN.png
- `cmd/zz_cmp/export_office.ps1`：PowerPoint COM 参考导出（-KeepAlive 可与其他自动化客户端共存）
- `cmd/zz_cmp`：逐页像素差 + 记录最差格 + HTML 报表
- `cmd/zz_roundtrip`：读取→写回，供生成侧对比
- 驱动与语料清单在 F:/pptx_corpus/_meta（run_compare2.sh、download_corpus.py、ordered_decks.txt）

## 已知剩余差距（按影响排序）

1. **SmartArt 布局引擎缺失（语境最大项）**：无缓存 drawingN.xml 的 SmartArt
   ——LibreOffice 系文稿 17 份——渲染占位符 vs Office 自动重排，约 17–46%
   2/页。PowerPoint 用 layoutN.xml 的算法从 data1.xml 现场排布；实现有限
   核心布局（单流向流程图、循环、层级）可摘掉这批。
2. **treck 主题装饰的父级位置**：custgeom-nofill-path 版式标记两个渲染都
   画材黄区，但位置与 Office 导出不一致（35.9%）。属 z 序/版式几何链探索。
3. **ooxtheme**：版式标题字号走母版 txStyles 的链路（15.5%，标题 44pt vs 28pt）。
4. **keyframes/ArtisticEffectSample**：动画帧与艺术效果（双色调如 duotone 已修）。
5. 文字反锯齿级差异存在于每个文字页（zz_cmp 的容差判定层面，结构性不可归零）。
