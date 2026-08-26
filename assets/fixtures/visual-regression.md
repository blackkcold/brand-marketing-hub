# vivo 视觉回归合成夹具
<!-- deck: 覆盖 md2pptx_vivo.py 的核心版式与边界场景 -->
<!-- deck: meta: Synthetic QA｜2026-08｜Internal test only -->

## P1｜封面
@layout bullets
- 合成品牌营销演示，不含真实项目、人物或版权素材
- 用于结构与布局 smoke test，不作为正式样张
> 封面备注：显式 P1 封面应避免额外生成重复封面。

## P2｜目录
@layout toc
- 1 基础版式覆盖
- 2 图文与表格覆盖
- 3 边界与校验覆盖

## P3｜Part 1 基础版式
@layout part
@sub 章节页用于确认 Part 数字、标题与副标题位置

## P4｜要点页：新品传播合成策略
@layout bullets
@sub 蓝色标题条与右侧结论句
- 目标：用完全合成内容验证标题条、正文行距和两级项目符号
  - 子项：二级缩进应保持可读
- 节奏：预热、发布、长尾三段互相承接
- 资产：短视频、门店海报、社群贴片均为占位描述
> 要点页备注：备注应写入 speaker notes，不进入可见画布。

## P5｜三卡片：渠道分工
@layout cards 3
### 自有阵地
- 官网与会员中心承接核心信息
- 社群推送保持统一口径
### 内容平台
- 合成短视频脚本突出体验点
- 评论区 FAQ 复用统一答疑
### 线下触点
- 门店屏幕播放核心卖点
- 导购话术保持简短一致

## P6｜四卡片：创意方向
@layout cards 4
### 方向 A
- 轻量开箱
- 场景第一视角
### 方向 B
- 夜景样张
- 朋友聚会
### 方向 C
- 长续航一天
- 通勤碎片化
### 方向 D
- 服务权益
- 以旧换新

## P7｜左图右文：体验场景
@layout split left
@img __VISUAL_FIXTURE_IMAGE__ | 合成渐变图片，仅用于 smoke test
- 画面：左侧图片应 cover-crop 填满区域
- 文案：右侧 bullet 区域不应与图注重叠
- 备注：图片由测试脚本在临时目录生成

## P8｜右图左文：门店物料
@layout split right
@img __VISUAL_FIXTURE_IMAGE__ | 合成门店灯箱示意
- 画面：右侧图片应保持裁切逻辑
- 文案：左侧信息流保持 3 条以内
- 校验：split left 与 split right 都应被解析保留

## P9｜表格：排期与负责人
@layout table
| 阶段 | 时间 | 产出 | 负责人 |
|---|---|---|---|
| 预热 | T-14 至 T-7 | 合成悬念海报 | 内容组 |
| 发布 | T 日 | 主视觉与核心卖点页 | 项目组 |
| 长尾 | T+1 至 T+21 | FAQ 与复盘材料 | 运营组 |

## P10｜时间线与图片指导
@layout bullets
@sub 后续人工截图检查的观察点
- 时间线：检查阶段条目是否在同一阅读顺序内呈现
- 图片指导：真实业务图片应替换为授权素材，本夹具只使用脚本生成图
- 截图建议：关注封面渐变、标题条、卡片间距、表格行高和图片裁切

## P11｜长 CJK 文本压力页
@layout bullets
- 长文本：这是一段完全合成的中文压力测试文本，用于观察较长 CJK 句子在固定宽度文本框内的自动换行、行距估算和潜在溢出提示，内容不指向任何真实产品或商业计划，只描述视觉回归测试所需的边界条件。
- 混排：vivo visual regression smoke fixture combines English tokens, numbers 12345, and 中文字符 to exercise font fallback.
- 收敛：该页允许触发溢出告警，但不应导致渲染失败。

## P12｜缺失图片降级
@layout split left
@img assets/fixtures/nonexistent-synthetic-image.png | 故意缺失：应降级为全宽文字区
- 缺失图片：renderer 当前逻辑应跳过图片并保留文字内容
- 目标：后续人工检查可确认无异常占位崩溃
- 约束：测试不得创建这个缺失文件

## P13｜占位符校验页
@layout bullets
- 原始夹具包含占位符用于验证 validator 失败路径：【待替换合成字段】
- smoke test 会在临时副本中替换该占位符，以验证通过路径

## P14｜备注与普通内容
@layout bullets
- 备注页用于确认 speaker notes 在 pptx 中可被写入
- 页面内容保持简短，避免与备注混淆
> 备注校验：这条 synthetic note 只用于结构检查。

## P15｜Thank you
@layout end
> 结束页备注：end layout 不应额外绘制页脚页码。
