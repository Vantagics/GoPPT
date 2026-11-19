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
| r115 | 0fb813c | pList 族（垂直项目符号列表）bullet 列：容器带（45% 高、
  75% 白混）+ 每列空框（68% 白混）+ 三张深色圆角卡（父标题+子项 bullet，autofit
  驱动缩字）；统一 quick style 用 deck 主题 accent1（解析时盖章进模型） | smartart-autofit-sync 31.66→0.66 |
| r116 | 17e60a4 | list1（垂直盒列表）描边行：全宽细边框白盒（63% 节距）
  + 左上叠实心圆角卡（69% 宽、5.4% 左移、斜体小标签） | smartart-linear-rule-vert 23.84→3.22 |
| r117 | f0aaf55 | 缓存 drawing 的 blipFill 按drawingN.xml.rels 解析（原误用 slide
  rels——rId1 撞车，照片圆读到 layout XML）；hier 族标签改自适应换行缩字 | smartart-simple 13.62→6.69 |
| r118 | da71341 | hier 引擎：叶列游标加 frame 的 x 偏移（原树整体左移出框）；
  asst 节点左挂独立层（deep 计数加带、子树下移一行、水平肘连主干）；卡宽上限 30% | org-chart 18.55→15.68、org-chart2 19.69→17.64 |
| r119 | 1e0fa40 | 多图 deck 的 frame 部件隔离：stash 按 dgm:relIds 四元组
  （dm/lo/qs/cs）过滤（原全 slide 部件入每个 frame，auxPaths 被最后一个 data
  覆盖——三 frame 全读第三图）；单节点竖带内缩 15%→4% | smartart-font-size 16.80→16.19 |
| r120 | 9d20ea1 | hProcess3 箭头：凹口轮廓（头部背上凸 18%、尖点 72% 高、
  体右下角收 92.5%）+ 打包进 frame 的 19%..63% 段；标签按体段均布 | smartart-linear-rule 11.25→2.52 |
| r121 | a39abd8 | snake 的 grDir="tR" 分支：data 无 dir 属性且 layout 带 tR 支
  时右到左镜像列序，网格缩至 72% 帧宽居中，尾盒按镜像前列号居中 | rightoleft 10.47→2.62、font-size 16.19→14.97 |
| r122 | b859ac6 | cycle 环校准（圈径 28% 短边、环 35%、标签随圈缩放、圈间
  顺时针浅色箭头）+ bullet 路径 wrapRunLineWithIndent 补逐字断行（88 字符
  长串溢出 slide） | cycle 9.74→5.15、autoTxRot 7.81→5.91、font-scale 9.35→8.47 |
| r123 | c87adc1 | accent-process 两盒对按导出实测：对宽 43%/间隙 14%、头卡
  8%..42% 高×82% 对宽左对齐、体卡半透明白+accent 边框右下叠 18%、体左上
  深色 bullet 行、间隙浅色箭头 | smartart-accent-process 10.20→5.69 |
| r124 | ca907b1 | 同位渐变 stop = 硬台阶（撤销 r100 的展开——之前把 Go/Office
  看反了：平滑的是我们、阶梯才是 Office）；渐变像素三处改 alpha-over 画布
  合成（80% alpha 暗 stop 呈现为底色上的阴影而非生 RGB） | minimal-gradient 9.29→0.47 |
| r125 | 9b21678 | 占位符梯级携带 cap="all"（layout lstStyle → Font.CapAll，
  渲染期大写文本——换行测量与绘制同字形，解析文本与 gen 往返不变）；
  梯级的 nil-font 跳过与无继承早退两处守卫为 cap-only 定义放行 | ArtisticEffectSample 10.03→9.93 |
| r126 | 887747b | bgRef 改解析 <a:bgFillStyleLst>（原误用 shape fillStyleLst——
  两列表每主题都不同）；blipFill 条目退化为 phClr solid；path/fillToRect/
  tileRect 构建真 FillGradientPath；母版 titleStyle 梯级也携带 cap="all" | themes 8.91→3.60（max 28.86→24.36） |
| r127 | b1ae01c | 组织架构图悬挂分支：节点最后一个带子树的子分支整 subtree
  垂直堆叠（运行行号递归、左缘主干+肘线、占一列）——Office 的经典悬挂律 | org-chart2 17.64→14.88 |
| r128 | eccdd2c | algn="just" 两端对齐：非末行按词间空格分摊行内富余
  （段内逐 run 分段绘制、粗体/下划线/删除线跨延伸距） | pptx2svg 8.28→8.29（对齐本身已一致，残差为换行点字体度量差） |
| r129 | c5377c2 | SmartArt 点文本保留 <a:br/> 为硬换行（LO 用 br 断行而非
  段落——收集器跨过去拼成一行）；smartArtWrap 按段落分段换行、cycle 标签
  与两盒 bullet 消费处净化为空格 | smartart-font-size 14.97→14.84 |
| r130 | bf64b2f | 母版 titleStyle 梯级读取渐变文字色（首 stop 作为级别
  solid 色——KEY02 的 48pt 白字写在图上，原因 solidFill-only 门控掉黑） | KEY02 6.12→5.76 |
| r134 | 35bc837 | <a:uFill>（LO 下划线填充）不再读作 run 文字色——LIBRE_OFFICE
  表格的显式黑字被 cell 的白 uFill 覆盖成白字 | LIBRE_OFFICE 6.16→6.09 |
| r135 | 41bbc73 | 表格样式带色文本不再覆盖显式命色的 run（runColorSet 守卫） | 显式色保真（无指标回归） |
| r138 | e2ef335 | 装饰形状 deco 标记：slide 部件只发自有形状、layout 部件携带
  自身装饰（deco 图片媒体旁路发射+layout rels 接线） | 21 deck gen 改善
  （smartart 族 16.8-37.4→0.2-0.9、ArtisticEffectSample 10.0→4.6、pptx2svg
  8.3→3.2）；4 deck 因单 layout 归属轻微回退 |
| r139 | d7ad48e | layout 装饰发射 groundwork：PlaceholderShape case 补齐；
  KEY02 gen 仍 96.63（deco pic 数据解析或 PlaceholderShape 路径需深查） | 改善基础设施落地 |
| r140 | ddebe48 | **补上丢失的 layoutPicXML 行**（r139 的发射循环写了媒体与 rels
  但 continue 跳过了 decoXML.WriteString(w.layoutPicXML(...))——装饰图片
  XML 从未进入 layout spTree） | KEY02 gen 96.34→**4.59**、2411→16.91 |
| r141 | 67a837e | SmartArt `<dgm:bg>` 背景色渲染：parseSmartArtBgColor 提取 +
  cached-drawing 前置填充 rect / model-only 走 BgColor | background-drawingml-fallback 7.94→**1.30** |
| r137 | 35f43e8 | schemeClr run 色分支补 runColorSet（原 srgbClr 有而 schemeClr 无——
  customGeo slide05 标题被 layout 黄色 defRPr 覆盖） | customGeo 7.01→7.00（标题白 ✓） |
| r131 | a86df4b | 占位符装饰保留填充/边框、清空提示文本（原规则整形状
  保留把 Click to edit... 提示叠画在幻灯片标题上；treck 图片标记的
  custGeom 绘制例外不变） | ShapeLineProperties 6.46→5.42 |
| r132 | 0987c70 | 椭圆渐变直接在画布上 alpha-over 合成（原在透明缓冲上
  混——75% alpha 的紫 stop 印成暗色；smartart-simple 的 def 圆透过
  重叠区读出底下的照片） | smartart-simple 6.85→5.66 |

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

1. **SmartArt composite 几何**：org-chart 15.7% — office 的 composite
   溢出声明 frame 重排（盒尺寸/行距），v1 引擎按 frame 网格布局；
   accent-process 5.7%、org-chart2 14.9%（悬挂律已实现）。
   composite 约束求解（r110 两盒对已收大头）。
2. **multi-master writer**：writer 创建一个 master/layout/theme；多母版
   文稿（themes 20.8% gen）的 master 2-7 需写回支持。
3. **customGeo 其余页**：最差页自 gdLst/pie/椭圆修复后 14.73%
   （slide28 组合页），已转细节样式。
5. **LO 布局模板内容**：font-size 的第二张 auto 卡等 LO 布局自带
   内容 v1 不复绘（r119 已让各 frame 各读各的 data）。
6. **ArtisticEffect/pptx2svg/accent-process**（8-10% 级）：艺术效果、
   SVG 图形细节与 accent-process 的 composite 重叠；smartart-simple 残差
   为 def 圆文字底色与重叠区像素。
6b. **smartart-rotation 6.87**：径向汇聚大箭头（块状杆+张开头部）已实现
   但标签需随箭头旋转——引擎无任意角度文字渲染，圆环近似按指标更优故
   回退（r132 实验）。
6d. **SmartArt 标签字号（autofit 模型）**：layout 的 ruleLst 声明
   primFontSz val=5 起步 + 无上限增长规则，且边距以 refType=primFontSz
   fact=0.3 随字号增长（负反馈）。两次实验均净负：整体提上限 65
   （font-size +0.61 但 6 deck 共损 2.7pp）、增长+0.6sz/0.2sz 边距反馈
   近似（同样净负）——需要把 constraint/rule 求解（含边距-字号反馈）
   移植进 fit 才能收敛。
6g. **r145 多 master/多 theme writer（themes gen 19.12→4.32，slide4 75→1.9）**：
   themes deck 是 7 master/7 theme（每 master 一个主题）。writer 现按
   masterPlan 发射每 master 部件：clrMap（master 可把 bg1 映到 dk1，Trek）、
   自有 bg（bgRef 1002 解析出的 Austin 渐变）、母版装饰形状（Rectangle 65
   白卡等）、sldLayoutIdLst 列自己的 layout；layout rels 指向各自 master；
   presentation.xml sldMasterIdLst + rels + contentTypes 全接线；theme 部件
   按读取时缓存的原始字节逐字回写（Austin/Trek/Tradeshow 的 clrScheme/
   fontScheme/fmtScheme 不再退化为硬编码 Office 主题）。占位 deco 以真
   <p:ph> 占位符回写（Office 只在 slide 实例化时绘制，slide4 的白色标题卡
   由此继承）。layout 根发 showMasterSp="0"（"layout turns off graphics"
   slide，51→2.6）。
   **0x80070570 修复要点**：sldMasterId 与 sldLayoutId 的 id 必须全包唯一——
   layout 段占 2147483648+idx，master 段放 2147483648+1000+idx（r133 的
   失败根源即此，非 OPC 接线）。python-pptx 可开、PowerPoint COM 严格
   Open 通过（designs=4）。
   10 个 gen deck 复测零回退。剩余 slide8 24%（Trek 贴图 tile）。
6p. **r153 行距 pin 实验组（负结果，全量回退）**：显式 <a:lnSpc><a:spcPct
   val="100000"/> 能把 alterman 的 body 行距钉回 1.2em（58→64px，与参考
   一致，7.59→5.36）——但同一 pin 在 KEY02（bodyStyle 声明了 90% 于
   lvl1-5、lvl6-8 未声明）与 master-slides（表全静默）上回退：这些
   deck 的原始渲染走字体度量行距，显式 100% 反而改变它们。三个变体
   （无差别 pin / 表级门控 / 仅未声明级别）都无法同时满足两族——
   Office 给"未声明行距"的段落哪种规则是包级/文档级判定，单件拼装
   无法复现。保留 lnSpcPct 捕获（master txStyles 的 lnSpc 现在进
   masterLevelStyle，梯子可推理），pin 全回退；各 deck 恢复 r152 值
   （alterman 7.59、KEY02 4.40）。master-slides 的"2.01"是陈旧导出
   对比的假值，真值 4.62（文字宽度漂移家族）。
6o. **r152 主题贴图四件套（themes gen 3.85→1.58，slide8 24.1→1.54）**：
   (1) 读取：theme 自带 rels 的图片按 blip r:embed 关系 id 缓存进
   themeState.media；(2) 主题解析：bgFillStyleLst 的 blipFill 不再退化为
   纯色——捕获 r:embed、duotone 的两组 schemeClr+ops、tile 标记；
   (3) bgRef 解析：blip 样式产出一个背景 DrawingShape（贴图数据 +
   duotone 两色（phClr 经 ops 解析）+ tile），走渲染端图片路径（支持
   duotone/tile）；(4) 写出：master 的 bgImage 发 bgPr blipFill（duotone
   内嵌 + tile + 媒体部件 + master rels），slide 的背景 DrawingShape
   同样回发 duotone（此前 slide bg writer 丢 duotone——slide9 曾被
   渲染成近白纹理）。L47 的 overrideClrMapping（bg2→dk2）经既有梯子
   生效：slide9 暗棕、slide8 浅沙各自正确。
   残余：minimal-gradient-fill-issue 0.19→1.27（标签文字像素漂移，
   已知项 7 家族，非贴图回归）。
6n. **r151 master 占位符逐字回写 + compatLnSpc 走梯 + fontNameSet
   （master-slides 3.94→2.01）**：master 的占位符 <p:sp>（继承定义——
   几何/bodyPr/每级 lstStyle）逐字缓存回写进 master 部件（deco 过滤器
   曾丢光无画占位符）；compatLnSpc="1"（legacy 1.2em 行距）沿
   占位符梯子传播到 slide 占位符并在 bodyPr 回发；fontNameSet 标记
   记录运行自述的字体——写侧不再把库默认 Calibri 发成显式字体
   （显式字体会顶掉主题 +mn-lt 并改变行距度量）。
   **alterman 残余 7.59（未破）**：body 行距 58px vs 参考 64px（24pt
   Arial：1.0875em vs 1.2em）——二分矩阵：源 txBody/主题/布局/
   defaultTextStyle 逐个换入我们的包均 58；我们的 slide20 放进源包
   即 64；master body ph/txStyles 已逐字一致、compatLnSpc 已回发仍 58。
   毒源在包级剩余差异（master 的 OLE graphicFrame vs pics、spTree
   顺序、bg schemeClr-vs-srgb 或 Office 的包级兼容模式判定）——
   待专轮。
6m. **r144–r150 全量 gen 普测（stdout 口径，cmp.log 是历史脚本遗留、
   勿再当现状读）**：67 deck 均值 1.13%。榜首：alterman 7.65（7px 文字
   恒定偏移，见 6l(a)）、tdf140865Wordart3D 7.63、LIBRE 5.55、KEY02 4.40、
   ShapeLineProperties 4.16、master-slides 3.94、table-with-theme 3.92、
   themes 3.85、2411 3.83、group 3.00。今日已清零/大幅收敛：
   themes 22.76→3.85、custgeom-nofill-path 35.90→0.00、
   tdf149588 6.33→0.76、minimal-gradient 13.06→0.19、fill-color-list→0、
   ooxtheme 6.44→0.61、group-rot→0.66、sample_pptx_grouping_issues→0、
   alterman 9.54→7.65、ArtisticEffectSample gen 3.66→1.11。
6l. **下一批归因（待做）**：(a) alterman 的残余 ~2.5%/张是文字整体
   上移 7px（1200px 渲染高）的恒定偏移——写出几何与源逐 EMU 相同，
   属渲染端字体度量/像素取整（已知项 7 的形态）。(b) themes slide8
   24%：Trek bgRef 1003 = theme bgFillStyleLst[2] 的 duotone blipFill
   贴图，媒体在 theme 自己的 rels（theme/_rels/themeN.xml.rels）——
   需要 theme 媒体缓存 + bgRef→FillPicture 解析 + 渲染端贴图 duotone
   + master bg blipFill 回写四件套。(c) tdf140865Wordart3D 7.63 =
   prstTxWarp/sp3d（roadmap 9）。
6k. **r149 线条填充的 <a:alpha>（tdf149588 6.33→0.76）**：tdf149588 的
   弧线是 noFill + 透明线条（ln solidFill schemeClr accent1 lumMod60/
   lumOff40 + alpha=0）——参考渲染里弧线完全不可见。我们的渲染把线条
   画成不透明粉/蓝（accent1 60% 亮度）。r144 的 alpha 修复只覆盖了
   writeFillXML 与渐变 stop；线条边框写出（writeBorderXMLWithEnds、
   DrawingShape 的 lnXML）仍用裸 colorRGB。改为 srgbClrXML。
6j. **r148 组的 chOff/chExt 逐字回写（alterman 9.52→7.65，group-rot 4.21→0.66，
   sample_pptx_grouping_issues→0.00）**：writer 把组的 off/ext 同时发到
   chOff/chExt——而读取侧给渲染解析了子形状坐标后仍保留原 child space。
   源组的 child space 常与组尺寸不同（alterman 的组把 1920x2496 的
   子空间放大 1587x），恒等 chOff/chExt 把每个子形状映射成左上角的
   一枚小点——组的填充形状全部消失（白框）。现在 chOff/chExt 发
   g.childOff/childExt（childExt=0 时回退 off/ext，覆盖 writer 自建组）。
   6 组系 deck 复测：group-rot 0.66、sample_pptx_grouping_issues 0.00、
   group-connector-descr 0.37、group-shadow 1.68，无回退。
6i. **r147 占位符 idx 匹配分档 + custGeom 逐子路径 fill（custgeom-nofill-path
   35.9→0.00 全对）**：两条 Office 法则落地。(a) layout 档的占位符匹配
   只认 idx——slide 的 pic idx=3 不继承 layout 的 pic idx=8 的 custGeom
   填充（曾把 9x5m 黄块画到 slide 上）；master 档保留 type 回退
   （00022693 slide17 的 idx=7 body 仍继承 master body 几何）。
   matchPlaceholderDef 加 typeFallback 参数。(b) 带画的非文字占位符
   （pic custGeom 黄块）Office 当普通形状渲染——writer 分文字/非文字：
   文字占位符回写真 <p:ph>，非文字回写普通形状。(c) custGeom 写出按
   子路径拆 <a:path>，子路径 fill="none" 逐字回写（读取侧 subPathNoFill
   已存，写出侧一直合并成一个 path 把透明区也填了）。
   TestReaderKeepsStandInForUnreadableChart 修随机失败（firstPartMatching
   走 map 顺序不定，r142 加 rels 后一半概率删错件；改为全删）。
6h. **r146 master txStyles 逐字回写 + 运行级 cap="all"（slide10 4.6→3.3）**：
   读取时暂存各 master 的原始 <p:txStyles>（title/body/other 的 lvl1-9
   阶梯，Trek titleStyle cap="all"），writer 逐字发射回 master 部件；
   占位符阶梯解析出的 font.CapAll 在运行 rPr 里发射 cap="all"——Go 渲染
   大写、写出文件里却没有声明，Office 渲染字面大小写（slide8 标题）。
   slide8 剩余：Trek bgRef 1003 是 theme bgFillStyleLst[2] 的 duotone
   blipFill 贴图（媒体在 theme 自己的 rels 里）——需要 theme 媒体部件
   + 贴图渲染 + bg blipFill 回写三件套（新特性，Roadmap）。
6f. **r144 多 layout writer 接线 + 填充 alpha（themes gen 20.65→19.12）**：
   之前 slide rels 硬编码指向 slideLayout1——多 layout deck（themes 的
   8 个 layout）写出的 layout 部件全是孤儿，slide 全部落在空 layout1 上。
   修：slide rels 按布局计划指向各自 layout；contentTypes 声明每个
   layoutN；master sldLayoutIdLst 列全（缺声明=PowerPoint 报修复）。
   layout 的 <p:bg><p:bgPr> 填充按 layout 部件发射（layoutBgCache）。
   三个机制修复：
   (a) 占位符 deco 不写入 layout——Office 只渲染被 slide 实例化的
       layout 占位符（themes L23 43 个 ph 形状写成普通形状后 slide7
       1.3%→92%）；paint 由读取侧 ladder 解析进 slide 自有占位符。
   (b) solid/gradient 填充发 <a:alpha>——ARGB 首字节 < FF 时（L23 的
       10% 白带 19FFFFFF 曾写成不透明白，slide6 56→2.6）。
   (c) slide 根发 showMasterSp="0"（"overridden background" slide 靠它
       压制 layout 图形，slide7 92→1.3）。
   10 个 gen deck 深对比（多 layout/透明填充/deco/bg）零回退；
   全量测试过。剩余：slide4/5 75% 属多主题（每 master 一个 clrScheme，
   writer 硬编码 Office 主题）= 6c 多 master 工程。
6e. **gen 回归（KEY02 96、2411 -14.7，r109 起）部分修（r138/r139）**：装饰形状
   被并进 slide.shapes 后 writer 把它们当 slide 内容重复发射——布局自身
   画布几何被改挂到 slide 上。deco 标记 + writer 过滤 + layout 部件携带
   自身装饰（含媒体与 rels）。剩余：多 layout deck 的装饰归属其首个
   layout（2411 +14.7、themes +6.9——writer 单 layout 模型的限制）。
6c. **multi-master writer（r133 实验）**：发射已通到 OPC 有效（python-pptx
   可开、关系目标齐全、layout id 全局唯一），但 PowerPoint 仍报
   0x80070570 拒开——包接线之外的 OOXML 一致性（母版样式化背景、
   notesMaster 一致性、每 master 主题保真）是独立工程。读取侧已为
   其备好每 slide 的 rung 路径戳记（r133 落地保留）。
7. **文字反锯齿与换行点**：结构性差异，不可归零（字体度量致个别
   词掉行，行高随之漂移——customGeo slide28 的表格行高漂移属此类）。
8. **writer 图表写出完整性**：radar/scatter chart RT-FAIL——r142 已补
   date1904/lang/roundedCorners/txPr + chart rels 发射；剩余需
   externalData 引用+嵌入 Excel（reader-writer 全链路新特性）。
9. **文字弯曲+3D 挤出**：tdf140865Wordart3D 7.62——prstTxWarp + sp3d
   的逐字符曲线定位与 3D 挤出渲染。
10. **图像艺术效果**：ArtisticEffectSample 9.93——a14:imgProps 的
   模糊/玻璃等图像滤镜渲染。
