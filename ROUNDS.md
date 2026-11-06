# 保真度轮次笔记（r75–r106）

从网上取得 79 个复杂度各异的 PPTX（LibreOffice core 回归集、apache/poi 实拍文稿、
POI 特性文稿），用 cmd/zz_render + cmd/zz_cmp 与 PowerPoint COM 导出
（cmd/zz_cmp/export_office.ps1）逐像素对比，双阶段：渲染保真与生成保真
（cmd/zz_roundtrip：读入→写出→Office 导出对比 Office 导出原稿）。
每轮归因到具体机制、修复并带回归测试提交。

## 修复轮次

| 轮 | 提交 | 修复 | 效果（vs Office 导出） |
|----|------|------|--------|
| r75 | 9180bb8 | 截断 zip 抢救（EOCD 缺失时按本地头扫描重建）+ 画布高度取整 | Divino_Revelado 12/12 页可读 |
| r76 | 6706d0b | 母版/版式非占位符装饰形状渲染 | alterman_security 16.4% → <8% |
| r77 | c95a8c3 | 每母版一套主题 + showMasterSp="0" 抑制 | 七母版七主题的 themes 导向正确 |
| r78 | dff001f | 显式 b="0" 压过母版 bold（boldSet 门控全梯级） | ShapeLineProperties 粗细/换行一致 |
| r79 | 613a146 | 图表自动标题 / 右侧图例档位与垂直居中 / 自动饼图半径 | pie-chart 8.6% → 4.2% |
| r80 | f135b1f | 母版 clrMap 重映射颜色名 | master-bg-color / onemaster-twolayouts → 0.00% |
| r81 | 30257bb | SmartArt 部件往返保留 | fill-color-list 生成侧 47.6% → 7.2% |
| r82 | 942c61a | 版式 overrideClrMapping（克隆映射表） | themes slide09 97.8% → 28.8% |
| r83 | 2b40e12 | 背景 blipFill `<a:tile/>` 平铺（重采样贴图块） | page_transparent_bitmap 22.1% → 0.95% |
| r84 | 6384348 | 图片占位符继承版式几何 | customshape-bitmapfill-srcrect 26.2% → 0.29% |
| r85 | e5fbea1 | custGeom 多子路径 fill 属性 + SmartArt 框架 rId 解析 | 损坏包修复可开 |
| r87 | 69694a6 | custGeom 子路径**追加**而非覆盖 | custgeom-nofill-path 35.9% → 0.05% |
| r88 | f542fbf | 较近梯级的字号/颜色压过较远 + 占位符 lumMod/lumOff | ooxtheme 标题 54pt/副标题灰正确 |
| r89/r90 | f114a0b | SmartArt v1 引擎（lin/snake/cycle/hier 四族 + 模型解析） | font-size 45.9→15.5% |
| r91 | — | hier 整洁树 + 肘形连接线 | org-chart2 结构对齐 |
| r92 | 218f336 | 文本 run/段落连接 + transit 穿越 | font-size 文本完整 |
| r93 | d83b25a | 节点标签 autofit（测宽换行 + 字面缩放） | 引擎标签匹配 Office autofit |
| r94–r95 | 371e21e | snake 网格 gutter + transit 递归 | rightoleft 19.0→10.5% |
| r96 | 1b9bfab | accent-process 两色渲染 | node 配色匹配 |
| r97 | 7dad42b | loCatId "list" 垂直路由 + 行间隙 | linear-rule-vert 30.2→23.8% |
| r98 | 4add079 | 叶子展开 + 文本捕获清理 | autofit-sync 37.4→28.2% |
| r99 | 371e21e | snake 网格 gutter | rightoleft 19.0→10.5% |
| r100 | 4ad86c3 | 双 stop 同位置渐变展开 | minimal-gradient 13.2→9.3% |
| r101 | 2ba79fe | 超宽单词逐字符断行 | font-scale 文字不再溢出 |
| r102 | e49862e | vert tree 层级布局（Parent 头带 + Child 行） | table-list 37.7→2.77% |
| r103 | 60515a6 | SmartArt 引擎测试覆盖（lin/snake/cycle 渲染） | 引擎覆盖 0→52-100% |
| r105 | 4909c2c | FillPicture 图片填充（AutoShape blipFill） | smartart-simple 海洋照片圆 |
| r106 | 9be024a | 非 ASCII 子弹字符符号 face 解析 | ◆/▲/■ 子弹受益 |
| r108 | 7c7ba0c | tile 背景 bgPr blipFill 写回 + alphaModFix | page_transparent_bitmap gen 87.4→0.95% |
| r109 | 4c099c1 | 母版装饰形状写入母版 XML | alterman 装饰 gen 保留 |
| r110 | f9cda61 | accent-process 族 header+body 两盒对 | smartart-accent-process 31.8→10.2% |
| r111 | bea24bb | **custGeom gdLst guide 公式求值**（val/\*/+-/?:/pin/min/max/abs/sqrt/at2/sin/cos/tan/ss；
  path 无 w/h 时回退形状盒）+ RichText 自定义几何的阴影像素轮廓（noFill 走描边影子、填充走多边形影子） | customGeo 8.37→8.02、最差页 22.66→17.86；ooxtheme 14.94→5.75（母版楔形装饰） |
| r112 | 7108bfc | **pie 预设几何**（Centre 锚定楔形，adj1/adj2 起止角 + 21600000° 环绕；
  fill/扇形轮廓阴影/border 两半径+弧三处接线） | customGeo 8.02→7.14、最差页 17.86→15.59（三扇饼页） |
| r113 | d0838bc | 文本形状恢复 prstGeom：RichTextShape/PlaceholderShape 带 presetShape，
  椭圆预设按椭圆轮廓填/描边（渐变经 fillEllipseImage 逐像素椭圆裁剪） | customGeo 7.14→7.01、最差页 15.59→14.73；smartart-simple 16.93→13.62 |
| r114 | 3e52adb | hList1（水平项目符号列表）横排卡片：仅 hList1 走横向
  （10% 列距 + 8% 上下边距，实测自 fill-color-list 导出），hList2/3 保持竖排 | fill-color-list 30.26→1.86 |

PowerPoint 把粘贴的 SmartArt/图表转成的 "Freeform N" 全部走
`<a:gdLst>` 公式引用（`<a:pt x="f5"/>`），旧解析器 `strconv.ParseInt("f5")`
静默失败把所有点堆到 0——路径退化成空多边形，只有包围矩形阴影可见
（白盒变 255×0.65 的灰、浅蓝幽灵卡消失、肘形连接线变灰带）。r111 求值
guide 链（惰性求值+备忘录+环守卫），并让自定义几何的阴影沿几何轮廓
投射而非包围盒。

## 渲染端改善汇总

| 改善幅度 | 文稿 |
|---|---|
| **>20pp** | font-size 45.9→16.8、themes 44.2→8.9、autofit-sync 37.4→28.2、table-list 37.7→2.77 |
| **10-20pp** | linear-rule-vert 30.4→23.8、custgeom 35.9→0.05、linear-rule 18.0→11.25、rightoleft 19.0→10.5、ShapeLineProps 18.9→6.5、alterman 16.4→6.4 |
| **像素级** | custgeom 0.05%、bitmapfill 0.29%、tile 0.95%、clrMap 双文稿 0.00% |

## 生成端改善汇总

63/79 文稿 gen < 5%。writer 的 SmartArt 部件往返 < 1%、tile 背景
alphaModFix 写回（r108）、母版装饰形状写回（r109）。

## 与 Office COM 导出一致性

- **<1%**：表、图核心、母版背景映射、装饰形状、平铺贴图背景、图片占位符、
  SmartArt 往返、多母版主题链、clrMap/overrideClrMapping、showMasterSp
- 文字反锯齿与 Mitchell 重采样差异为结构性，不可归零

## 已知剩余差距

1. **SmartArt composite 几何**：accent-process 10.2% — 重叠位置的
   composite 约束求解（r110 两盒对已收大头）。
2. **autofit-sync vert tree 容器配色**：31.7% — office 的淡色圆角容器
   与 accent 底色头带的视觉差异需精确调参。
3. **multi-master writer**：writer 创建一个 master/layout/theme；多母版
   文稿（themes 20.8% gen）的 master 2-7 需写回支持。
4. **customGeo 其余页**：最差页自 gdLst/pie/椭圆修复后 14.73%
   （slide28 组合页），已转细节样式。
5. **radial 渐变**：minimal-gradient 9.3% — lumMod 在 gradient stop 的
   schemeClr 上的色彩变换链路。
6. **LO 布局模板内容**：font-size 的 "smaller" 盒等 LO 布局自带内容
   v1 不复绘。
7. **ArtisticEffect/pptx2svg**（8-10% 级）：艺术效果与 SVG
   图形细节。
8. **文字反锯齿**：结构性差异，不可归零。
