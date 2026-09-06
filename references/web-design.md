# 关系网 · 网页设计与调参说明

模板：`scripts/template.html`（零依赖单文件，纯手写物理引擎，不引 D3）

## 一、为什么不用固定布局

环形 / 树形布局是**摆拍**：位置是写死的，观众看到的是一张图，读不出关系。
力导向布局的位置是**算出来的**——谁跟谁紧、谁跟谁松，一眼可辨，
而且筛选掉一类关系后全网会重新聚簇，这个「重排」本身就是信息。

## 二、物理引擎四件套

### 1. 斥力：各向异性椭圆排斥（关键）

节点不是圆，是「圆 + 下方两行标签」的**高瘦矩形**。用圆做斥力一定会压字。

```js
// 建节点时量出真实占位
n.hw = max(n.r, 名字宽/2+5, 副标题宽/2+5);      // 半宽
const top = -(n.r+4), bot = n.r+16+sub字号+3;   // 竖直范围
n.hh = (bot-top)/2;  n.offY = (bot+top)/2;      // 半高 / 中心下移量
```

```js
const rx = a.hw+b.hw+30, ry = a.hh+b.hh+26;     // 排斥椭圆
const d  = Math.hypot(dx/rx, dy/ry);
if (d < 1.5) {                                   // 阈值 1.5
  const push = (1.5-d) * 3.4 * alpha;
  ...按真实空间单位向量推开...
}
```

`ry > rx` 是故意的：标签让竖直方向更占地方，所以竖直方向要推得更开。

### 2. 硬防重叠：deoverlap()

**必须独立于 alpha 存在。** 物理收敛后 alpha→0，斥力停了，
但只要有一处标签压字就毁了整张图。

```js
function deoverlap(){          // 每帧无条件跑 3 轮，直接改位置不加进速度
  const ox = (a.hw+b.hw+18) - Math.abs(dx);
  const oy = (a.hh+b.hh+13) - Math.abs(dy);
  if (ox<=0 || oy<=0) continue;
  if (oy < ox) 沿 y 推开   // 沿穿透较浅的轴分离，位移最小
  else         沿 x 推开
}
```

钉住的节点（`fx!==null`）不参与位移，只让对方让开。

### 3. 弹簧：按关系类型分档

不同关系的「紧密度」不同，长度和弹性都要分开给：

| 类型 | L（自然长度） | k（弹性） | 意图 |
|---|---|---|---|
| 情感 | 104 | .075 | 最紧、最强 |
| 血亲 | 118 | .055 | |
| 同盟 | 126 | .050 | |
| 仇杀 | 158 | .038 | |
| 控制 | 172 | .026 | 最松，虚线 |

度数高的节点要减权，否则枢纽人物会被四周扯成一团：
`权重 = 1/(1 + deg*0.25)`

### 4. 向心力：各向异性

屏幕上横向空间比纵向宽，让网横向铺开：

```js
n.vx -= n.x * 0.0009 * alpha;   // 横向松
n.vy -= n.y * 0.0030 * alpha;   // 纵向紧
```

## 三、启动序列（顺序错画面就会漂）

```js
alpha = 0.34; for (i<260) step0();       // 预热：定 alpha 推开，不衰减
for (i<420) { step0(); alpha *= .994; }  // 收敛到稳定形态
fitView();                                // 再按包围盒适配视野
alpha = 0.05; reset(); step();            // 只留余温，避免加载后继续漂
```

**教训**：早期版本预热后就交给实时模拟，结果 `fitView()` 算完视野，
物理还在以 alpha=.30 继续跑，画面又漂出去 180px。
必须**先收敛再适配**，之后只留极小的 alpha。

`step0()` 和 `step()` 用同一套力，区别只是 `step0` 不自己衰减 alpha（由调用方控制）。

## 四、fitView()

```js
const availW = max(320, W()-372);   // 右侧 372px 留给侧栏
const availH = max(260, H()-140);   // 上方 140px 留给标题
const s = min(1.4, availW/(spanX+56), availH/(spanY+56));
```

- scale 上限 **1.4**：再大节点字号失真
- 带 400ms 缓动的版本用于「筛选后重排」，让观众看清网络怎么变的
- 隔离模式、重置、加载后都要调一次

## 五、多重关系的双向弯曲

同一对人物常有两条关系（既仇杀又控制）。按出现次序给曲率：

```js
cur = (idx - (pairCount-1)/2) * 0.30;   // -0.15 / +0.15 对称弯开
```

控制点：`中点 + 垂直于连线的偏移(cur * L * 0.5)`

## 六、视觉规范

**主题**：默认深色（`--bg0:#0b0e14` → `--bg1:#151b26` 径向渐变）。
关系网类可视化深色更有质感，但必须提供浅色切换。

**文字可读性**：所有 SVG 文字加同底色描边，压在连线上也读得清：

```css
text { paint-order:stroke; stroke:var(--halo); stroke-linejoin:round; }
.lbl { stroke-width:3px }  .sub2 { stroke-width:2.4px }
```

`--halo` 深浅各一值（深 `#0f1520` / 浅 `#f4f7fb`）。
emoji 类文本（📌）要 `stroke:none`，否则描边糊成一团。

**阵营配色**：内置 8 组，每组深浅两套 `[填充, 描边, 文字]`。
深色用低饱和填充 + 亮描边，浅色用浅填充 + 深文字。

**层次**：ring=0（主角）独占——半径最大、文字用阵营文字色（其他人用全局文字色）、
加 glow 滤镜、常驻一圈 halo 光晕。

**连线**：虚线类型（控制/胁迫）在高亮时加 `stroke-dashoffset` 流动动画，
视觉上表达「这条线是活的、在施压」。

## 七、交互清单（别删项）

| 操作 | 实现要点 |
|---|---|
| 拖拽 | mousedown 在节点上要 `stopPropagation`，否则同时触发画布平移 |
| 钉住 | 松手时写 `fx/fy` 并显示 📌；双击节点解钉 |
| 缩放 | 以光标为锚点：`tx = clientX - rect.left - worldX * newScale` |
| 隔离 | 双击 → 只留一阶邻域 → 静默收敛 → fitView |
| 筛选 | 关系类型 × 阵营 × 隔离态三者取交集，统一由 `recalc()` 算 |
| 悬停边 | 细线难命中，叠一条 `stroke-width:16` 的透明 `path.hit` |
| 搜索 | 命中 1 人直接飞过去；命中多人侧栏列清单 |

## 八、验证方式（重要）

**本环境 chrome-devtools MCP 的 `take_screenshot` 写不进工作区**
（报 `Access denied: path is not within any of the configured workspace roots`），
所以**不能靠看图验收**，必须写程序化断言：

```js
// 1. 重叠数（必须为 0）
const b = nodes.filter(n=>!n.hide).map(n=>({l:n.x-n.hw, r:n.x+n.hw,
           t:n.y+n.offY-n.hh, b:n.y+n.offY+n.hh}));
// 两两求交，统计相交对数

// 2. 可见计数
document.getElementById('sN').textContent  // 节点数
document.getElementById('sE').textContent  // 边数

// 3. 屏幕包围盒是否落在可视区
nodes.map(n=>({x:n.x*scale+tx, y:n.y*scale+ty}))

// 4. 每个交互都 dispatch 一遍再断言
by.si.g.dispatchEvent(new MouseEvent('mouseenter'));
document.querySelector('.chip[data-t="ctrl"]').click();
document.getElementById('bReset').click();
```

**必测路径**：加载 → 悬停 → 逐个关关系类型 → 关阵营 → 隔离 → 退出隔离
→ 切主题 → 重置。每一步之后都查一遍重叠数和可见计数。
