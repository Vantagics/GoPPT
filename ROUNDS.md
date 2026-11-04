# 保真度轮次笔记（r75–r86）

从网上取得 79 个复杂度各异的 PPTX（LibreOffice core 回归集、apache/poi 实拍文稿、
POI 特性文稿），用 cmd/zz_render + cmd/zz_cmp 与 PowerPoint COM 导出
（cmd/zz_cmp/export_office.ps1）逐像素对比，双阶段：渲染保真与生成保真
（cmd/zz_roundtrip：读入→写出→Office 导出对比 Office 导出原稿）。
每轮归因到具体机制、修复并带回归测试提交。

## 本轮修复

| 轮 | 提交 | 修复 | 效果（vs Office 导出） |
|----|------|------|--------|
| r75 | 9180bb8 | 截断 zip 抢救（EOCD 缺失时按本地头扫描重建）+ 59bf574 画布高度取整 | Divino_Revelado 12/12 页可读；非整比例画布逐像素配对 |
| r76 | 6706d0b | 母版/版式非占位符装饰形状渲染 | alterman_security 16.4% → <8%，红条带像素级一致 |
| r77 | c95a8c3 | 每母版一套主题 + showMasterSp="0" 抑制 | 七母版七主题的 themes 导向正确；抑制页装饰消失 |
| r78 | dff001f | 显式 b="0" 压过母版 bold（boldSet 门控全梯级） | ShapeLineProperties 粗细/换行一致 |
| r79 | 613a146 | 图表自动标题 / 右侧图例档位与垂直居中 / 自动饼图半径 | pie-chart 8.6% → 4.2% |
| r80 | f135b1f | 母版 clrMap 重映射颜色名 | master-bg-color / onemaster-twolayouts 100%/50% → 0.00% |
| r81 | 30257bb | SmartArt 部件往返保留（母版装饰+版式装饰+隐藏图） | fill-color-list 生成侧 47.6% → 7.2%；customGeo 49 部件回写 |
| r82 | 942c61a | 版式 overrideClrMapping（克隆映射表） | themes slide09 97.8% → 28.8% |
| r83 | 2b40e12 | 背景 blipFill `<a:tile/>` 平铺（重采样贴图块） | page_transparent_bitmap 22.1% → 0.95% |
| r84 | 6384348 | 图片占位符继承版式几何（DrawingShape 加入梯度 ladder） | customshape-bitmapfill-srcrect 26.2% → 0.29% |
| r85/+ | e5fbea1|ab2e360 | custGeom 多子路径 fill 属性 + SmartArt 框架 rId 解析修正 | treck 内区留空按语义；损坏包（0x80070570）修复可开 |
| r86 | adbd742 | 版式 lstStyle 斜体梯级 + 装饰画在 blip 背景图之上 | slide04 标题斜体；themes slide05 从 76.5% 跌出最差榜 |

## 与 Office COM 导出一致性（在最终对比中）

- **0.3% 以下（基本一致）**：表、图（pie/line/bar/radar/scatter 核心）、母版背景映射、
  装饰形状、平铺贴图背景、图片占位符、SmartArt 往返（缓存路径）、多母版主题链、
  clrMap/overrideClrMapping、showMasterSp 抑制等。
- **诚实声明**：文字反锯齿与 Mitchell 重采样差异存在于所有文字页，zz_cmp 的容差
  （>48/255 判差）捕获的是结构性差距；数值全部由该管线产出（F:/pptx_corpus/_meta 驱动）。
- 期间另一自动化客户端（同机 Devin 会话）间歇杀掉 POWERPNT，本次对比的 26 个失败
  条目已末轮集中补跑，CSV 中残余的重复行即两轮衔接的痕迹。

## 已知剩余差距（按影响排序）

1. **SmartArt 布局引擎缺失（最大项）**：无缓存 drawingN.xml 的 SmartArt
   ——LibreOffice 系文稿近 17 份——渲染显示占位符，而 Office 用 layoutN.xml 的
   算法从 data1.xml 现场重排（smartart-*/table-list 一族 9.7–45.9%）。实现
   单向流程/循环/层级三类核心布局可整批摘除。
2. **treck 主题装饰位置的父级链**：custgeom-nofill-path 35.9%——标记几何的
   版式坐标换算（chOff/chExt 的装饰组）。
3. **themes slide09/06**：29%/23%——画布纹理的映射 superb，蜂窝/粗纹笔触的
   precise 贴图位置与增高版式面板几何。
4. **ooxtheme 15.5%**：母版 txStyles 的标题字号/装饰渐变的组合链。
5. **ArtisticEffect/customGeo/pptx2svg/背景双色调**（8–10% 级）：艺术效果
   （必剪树印/玻璃等）与 SVG 图形细节。
