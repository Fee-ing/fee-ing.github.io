/*
 * 共享的「鲜艳多彩流体」背景引擎。
 *
 * 参考页：https://www.color4bg.com/zh-hans/generator/vivid-colorful-fluid-background
 * 原实现：https://github.com/winterx/color4bg.js → AbstractBackground/AestheticFluidBg.js（MIT）
 *
 * ## 为什么放在 public/ 而不是 src/
 *
 * 同一份实现要服务三个生命周期完全不同的消费者：
 *   1. index.html 的首屏静态壳 —— React 与 bundle 都还没到，只能是 classic script；
 *   2. React 侧的 LoadingMask —— 启动遮罩 / 页面加载遮罩 / 忙碌遮罩；
 *   3. 桌面壁纸。
 * 打进 bundle 就只有（2）能用。放在 public 由 <script src> 直接加载（同
 * loading-mask.css 的思路），三者共用同一个全局实例。
 *
 * ## 单例：一个上下文 + 一条循环，与「显示几处」无关
 *
 * 引擎懒建、全局唯一，持有**一个 WebGL 上下文 + 一条 rAF 循环 + 一张只画一次的
 * 512×512 染料纹理**。每个消费者交出一张 <canvas> 就登记成一个 surface：
 *
 *   - GPU 侧只渲染一次到引擎私有（不在 DOM 里）的离屏画布，尺寸取所有 surface
 *     的最大值；每次渲染用 viewport 偏移把画面落到该画布**左上角**的 w×h 子区域
 *     —— 避免逐帧改画布尺寸（改尺寸会重新分配整块后备存储）；
 *   - 再把这块子区域 drawImage 到各 surface 自己的 2D 上下文。这是浏览器内部的
 *     纹理拷贝，比开 N 个 WebGL 上下文便宜得多（上下文是稀缺资源，浏览器通常
 *     只允许十几二十个，且每个都占独立的 GPU 显存）。
 *
 * 每个 surface 按自己的尺寸独立取景（各自的网格几何 + 正交视口 + zoom 规则），
 * 所以窗口内的遮罩与全屏壁纸画面都正确，而不是同一张图被拉伸。
 *
 * ## 显隐与生命周期
 *
 * - surface 的 canvas 一旦离开 DOM（React 接管静态壳、遮罩隐藏等）自动摘除；
 * - 全部摘除后 rAF 循环自行停止；document 被隐藏（切标签）时同样停摆；
 * - `coverAll` 的 surface（不透明盖满整屏，如启动遮罩）存在时，其余 surface
 *   当帧不再渲染 —— 反正看不见，省掉一次全屏拷贝；
 * - 监听 webglcontextlost / restored：驱动重置后自动重建全部 GL 资源。
 *
 * ## 降级
 *
 * 任何一步失败（无 WebGL、着色器编译失败、语法不被支持……）都只是让
 * window.LoadingFluid 不可用或 mount() 返回 null，消费方各自退回 CSS 皮肤。
 */
(function (global) {
  'use strict';

  /* ===================== 常量（与参考页逐位对齐） ===================== */

  /** 染料层分辨率。固定值，与 surface 尺寸无关（原库同此）。 */
  const RTT_SIZE = 512;

  /** 「Blur Scale」：波纹采样幅度（原实现 uMagnitude，参考页滑块默认 0.15）。 */
  const WAVE_SCALE = 0.15;

  /** 参考页使用的 6 色色板（vivid 主题下的「Electric Shock」）。 */
  const COLORS = ['#7CFFFD', '#1C75ED', '#383A3C', '#EDE82F', '#FCBD33', '#EC2025'];

  /**
   * 6 团色斑参数 [x, y, inner, outer]。
   *
   * 原实现用 seedrandom(seed=1000) 现算；种子固定、调用顺序固定，结果必然恒定，
   * 于是这里直接把「离线用同一套 seedrandom 跑出的 24 个随机数」内联成常量 ——
   * 省掉一份 RC4 生成器，同时与参考页逐像素一致（inner/outer 已按
   * radius_inner = 0.1、radius_outer = 0.3 换算完毕）。
   */
  const DYES = [
    [0.107553, 0.437435, 0.145832, 0.335723],
    [0.497775, 0.454916, 0.176795, 0.406987],
    [0.59616, 0.31328, 0.180302, 0.421516],
    [0.976991, 0.045565, 0.104522, 0.430416],
    [0.406615, 0.357722, 0.197768, 0.418944],
    [0.641485, 0.281848, 0.162784, 0.462049],
  ];

  /**
   * 两处顶点形变锚点：把半径 900 内、以锚点为中心的高斯权重顶点朝 moveTo 拉。
   *
   * 权重 = σ=0.25 的高斯 / 1.6，再减「5 / 离上下边缘的距离」（防止把画面边缘
   * 一起卷进来）。这层位移让平面局部被拉伸，与波纹采样叠加才像液体。
   */
  const MORPHS = [
    { anchorX: 240, anchorY: -200, anchorZ: 0, radius: 900, moveToX: 260, moveToY: 160, moveToZ: 400 },
    { anchorX: -240, anchorY: 200, anchorZ: 0, radius: 900, moveToX: -260, moveToY: -160, moveToZ: 400 },
  ];

  /** 网格细分：与原实现一致（99 段 → 100×100 顶点、58806 个索引）。 */
  const SEGMENTS = 99;
  const ROWS = SEGMENTS + 1;

  /** 正交相机裁剪面距离，与原实现一致。 */
  const NEAR = 0.1;
  const FAR = 10001;

  /** 相机在 +z 轴 8000 处俯视原点 → 模型视图矩阵即一个 (0, 0, -8000) 平移。 */
  const MODEL_VIEW = new Float32Array([
    1, 0, 0, 0,
    0, 1, 0, 0,
    0, 0, 1, 0,
    0, 0, -8000, 1,
  ]);

  /**
   * 上下文属性。
   *
   * 布尔项全部关掉（不做混合、不用深度/模板）—— 两层都是不透明全屏绘制，
   * 一个像素一个颜色，这些状态纯属多余开销。
   *
   * preserveDrawingBuffer 保持 false：离屏画布不在 DOM 里，浏览器永远不会
   * 「present」它，也就不存在合成后清空的问题；而每次 blit 都发生在同一个任务里
   * （drawElements 之后立刻 drawImage），读到的一定是刚画的这一帧。
   */
  const CONTEXT_ATTRS = {
    alpha: false,
    depth: false,
    stencil: false,
    antialias: false,
    premultipliedAlpha: false,
    preserveDrawingBuffer: false,
    powerPreference: 'default',
  };

  /* ===================== 着色器 ===================== */

  /** 染料层顶点着色器：全屏四边形，坐标直接写裁剪空间（片元只吃 gl_FragCoord）。 */
  const DYE_VERTEX = `
    attribute vec2 position;
    attribute vec2 uv;

    varying vec2 vUv;

    void main() {
      vUv = uv;
      gl_Position = vec4(position, 0.0, 1.0);
    }
  `;

  /** 染料层片元着色器：渐变底 + 6 团柔边色斑，逐层 mix。 */
  const DYE_FRAGMENT = `
    #ifdef GL_ES
    precision mediump float;
    #endif

    #define TWO_PI 6.28318530718

    uniform vec2 u_resolution;
    uniform float u_time;

    uniform vec3 u_color_0;
    uniform vec3 u_color_1;
    uniform vec3 u_color_2;
    uniform vec3 u_color_3;
    uniform vec3 u_color_4;
    uniform vec3 u_color_5;
    uniform vec4 u_dye_0;
    uniform vec4 u_dye_1;
    uniform vec4 u_dye_2;
    uniform vec4 u_dye_3;
    uniform vec4 u_dye_4;
    uniform vec4 u_dye_5;

    // 柔边色斑：inner 以内完全不透明，到 outer 衰减为 0
    vec4 blurDot(vec3 color, vec2 st, vec2 pos, float inner, float outer) {
      float pct = distance(st, pos);
      float alpha = 1. - smoothstep(inner, outer, pct);
      return vec4(color.rgb, alpha);
    }

    void main() {
      vec2 st = gl_FragCoord.xy / u_resolution;
      vec3 color = vec3(1.0);

      vec4 dot_0 = blurDot(u_color_0, st, u_dye_0.xy, u_dye_0[2], u_dye_0[3]);
      vec4 dot_1 = blurDot(u_color_1, st, u_dye_1.xy, u_dye_1[2], u_dye_1[3]);
      vec4 dot_2 = blurDot(u_color_2, st, u_dye_2.xy, u_dye_2[2], u_dye_2[3]);
      vec4 dot_3 = blurDot(u_color_3, st, u_dye_3.xy, u_dye_3[2], u_dye_3[3]);
      vec4 dot_4 = blurDot(u_color_4, st, u_dye_4.xy, u_dye_4[2], u_dye_4[3]);
      vec4 dot_5 = blurDot(u_color_5, st, u_dye_5.xy, u_dye_5[2], u_dye_5[3]);

      color = mix(u_color_0, u_color_1, st.x);
      color = mix(color, u_color_2, st.x * st.x + -0.040);

      color = mix(color, dot_0.rgb, dot_0.a);
      color = mix(color, dot_1.rgb, dot_1.a);
      color = mix(color, dot_2.rgb, dot_2.a);
      color = mix(color, dot_3.rgb, dot_3.a);
      color = mix(color, dot_4.rgb, dot_4.a);
      color = mix(color, dot_5.rgb, dot_5.a);

      gl_FragColor = vec4(color, 1.0);
    }
  `;

  /** 主层顶点着色器：标准 MVP。 */
  const PLANE_VERTEX = `
    attribute vec3 position;
    attribute vec2 uv;

    uniform mat4 modelViewMatrix;
    uniform mat4 projectionMatrix;
    varying vec2 vUv;

    void main() {
      vUv = uv;
      gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0);
    }
  `;

  /** 主层片元着色器：对染料纹理做正弦波纹偏移采样 → 流动感。 */
  const PLANE_FRAGMENT = `
    // 优先 highp（WebGL2 必定支持）；退回 WebGL1 的老设备上高精度不一定可用
    #ifdef GL_FRAGMENT_PRECISION_HIGH
    precision highp float;
    #else
    precision mediump float;
    #endif

    uniform sampler2D tMap;
    uniform float uTime;
    uniform float uMagnitude;

    varying vec2 vUv;

    const float speed = 15.0;

    void main() {
      vec2 wavyCoord;
      wavyCoord.s = vUv.s + (sin(uTime + vUv.t * speed) * uMagnitude);
      wavyCoord.t = vUv.t + (cos(uTime + vUv.s * speed) * uMagnitude);
      vec4 frameColor = texture2D(tMap, wavyCoord);
      gl_FragColor = frameColor;
    }
  `;

  /* ===================== 纯函数工具 ===================== */

  /**
   * 正交投影矩阵（列主序），与 OGL `Mat4.fromOrthogonal` 逐元素一致。
   *
   * @param {number} left 左裁剪面（已含 zoom）
   * @param {number} right 右裁剪面
   * @param {number} bottom 下裁剪面
   * @param {number} top 上裁剪面
   * @param {number} near 近裁剪面
   * @param {number} far 远裁剪面
   * @returns {Float32Array} 4×4 矩阵
   */
  function orthoMatrix(left, right, bottom, top, near, far) {
    const lr = 1 / (left - right);
    const bt = 1 / (bottom - top);
    const nf = 1 / (near - far);
    return new Float32Array([
      -2 * lr, 0, 0, 0,
      0, -2 * bt, 0, 0,
      0, 0, 2 * nf, 0,
      (left + right) * lr, (top + bottom) * bt, (far + near) * nf, 1,
    ]);
  }

  /**
   * 十六进制色值 → 0..1 的 RGB。
   *
   * 刻意不做 sRGB → 线性转换：色卡库的 OGL Color 也只除以 255，保持一致。
   *
   * @param {string} hex 形如 `#RRGGBB`
   * @returns {number[]} [r, g, b]
   */
  function hexToRgb(hex) {
    const value = parseInt(hex.slice(1), 16);
    return [((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255];
  }

  /**
   * 按原实现的 MorphSurfaceConvex 逐锚点处理顶点：先登记权重，再整体位移。
   * 两个锚点串行处理（后者的距离从「已被前者拉过的顶点」算起），与原实现一致。
   *
   * @param {Float32Array} positions 顶点位置（就地修改）
   */
  function applyMorph(positions) {
    MORPHS.forEach((morph) => {
      const weights = new Map();

      for (let i = 0; i < ROWS * ROWS; i += 1) {
        const ix = i % ROWS;
        const iy = (i - ix) / ROWS;
        // 只形变内部顶点：四周边圈保持原样，画面边缘才不会被卷起来
        if (ix === 0 || ix === SEGMENTS || iy === 0 || iy === SEGMENTS) continue;

        const x = positions[i * 3];
        const y = positions[i * 3 + 1];
        const z = positions[i * 3 + 2];
        const distance = Math.hypot(x - morph.anchorX, y - morph.anchorY, z - morph.anchorZ);
        if (distance >= morph.radius) continue;

        // 同列的首/末行顶点 → 换算「离上下边缘还有多远」，用于边缘处压权重
        const top = ix;
        const bottom = SEGMENTS * ROWS + ix;
        const toTop = Math.hypot(
          x - positions[top * 3], y - positions[top * 3 + 1], z - positions[top * 3 + 2]
        );
        const toBottom = Math.hypot(
          x - positions[bottom * 3], y - positions[bottom * 3 + 1], z - positions[bottom * 3 + 2]
        );
        const toEdge = Math.min(toTop, toBottom);

        const factor = distance / morph.radius;
        // 高斯权重（σ = 0.25，除以 1.6 归一）；3.14 而非 Math.PI，与原实现逐位对齐
        let weight =
          ((1 / (Math.sqrt(2 * 3.14) * 0.25)) * Math.exp(-(factor * factor) / (2 * 0.25 * 0.25))) / 1.6;
        weight -= 5 / toEdge;
        weights.set(i, weight);
      }

      const dx = morph.moveToX - morph.anchorX;
      const dy = morph.moveToY - morph.anchorY;
      const dz = morph.moveToZ - morph.anchorZ;

      weights.forEach((weight, i) => {
        if (weight <= 0.001) return;
        positions[i * 3] += dx * weight;
        positions[i * 3 + 1] += dy * weight;
        positions[i * 3 + 2] += dz * weight;
      });
    });
  }

  /**
   * 生成铺满「surface 较长边」的方形网格（顶点为中心在原点的世界坐标），
   * 并把形变一次性烘进顶点里。
   *
   * 每个 surface 只在首次拿到有效尺寸时调用一次：之后尺寸变化只更新投影
   * （沿用原库 ColorBg.resize 的 zoom 规则），不重建几何 —— 与原实现一致。
   *
   * @param {number} size 边长（surface 宽高最大值）
   * @returns {{positions: Float32Array, uvs: Float32Array, indices: Uint16Array}}
   */
  function buildPlane(size) {
    const vertexCount = ROWS * ROWS;
    const positions = new Float32Array(vertexCount * 3);
    const uvs = new Float32Array(vertexCount * 2);
    const indices = new Uint16Array(SEGMENTS * SEGMENTS * 6);
    const step = size / SEGMENTS;

    let vertex = 0;
    for (let iy = 0; iy <= SEGMENTS; iy += 1) {
      const y = size / 2 - iy * step;
      for (let ix = 0; ix <= SEGMENTS; ix += 1, vertex += 1) {
        positions[vertex * 3] = ix * step - size / 2;
        positions[vertex * 3 + 1] = y;
        positions[vertex * 3 + 2] = 0;
        uvs[vertex * 2] = ix / SEGMENTS;
        uvs[vertex * 2 + 1] = 1 - iy / SEGMENTS;
      }
    }

    let offset = 0;
    for (let iy = 0; iy < SEGMENTS; iy += 1) {
      for (let ix = 0; ix < SEGMENTS; ix += 1) {
        const a = iy * ROWS + ix;
        const b = (iy + 1) * ROWS + ix;
        const c = b + 1;
        const d = a + 1;
        indices[offset] = a;
        indices[offset + 1] = b;
        indices[offset + 2] = d;
        indices[offset + 3] = b;
        indices[offset + 4] = c;
        indices[offset + 5] = d;
        offset += 6;
      }
    }

    applyMorph(positions);
    return { positions, uvs, indices };
  }

  /* ===================== 引擎（单例） ===================== */

  /** @type {object|null|false} null=未尝试, false=不可用, object=就绪 */
  let engine = null;

  /**
   * 建立引擎。整个过程零全局副作用，也从不往 DOM 里插自己的画布。
   *
   * @returns {object|null} WebGL 不可用 / 着色器失败时返回 null
   */
  function createEngine() {
    const glCanvas = document.createElement('canvas');
    const gl = glCanvas.getContext('webgl2', CONTEXT_ATTRS) ||
      glCanvas.getContext('webgl', CONTEXT_ATTRS);
    if (!gl) return null;

    /** 会被上下文丢失摧毁的 GL 资源集中在这里，恢复时整体重建。 */
    const res = {
      dyeProgram: null,
      planeProgram: null,
      dyeTexture: null,
      dyeFramebuffer: null,
      quadBuffer: null,
      dyeAttribs: null,
      dyeUniforms: null,
      planeAttribs: null,
      planeUniforms: null,
    };

    let contextLost = false;

    const compile = (type, source) => {
      const shader = gl.createShader(type);
      gl.shaderSource(shader, source);
      gl.compileShader(shader);
      if (gl.getShaderParameter(shader, gl.COMPILE_STATUS)) return shader;
      console.error('[LoadingFluid] shader compile failed:', gl.getShaderInfoLog(shader));
      gl.deleteShader(shader);
      return null;
    };

    const link = (vertexSource, fragmentSource) => {
      const vertexShader = compile(gl.VERTEX_SHADER, vertexSource);
      const fragmentShader = compile(gl.FRAGMENT_SHADER, fragmentSource);
      if (!vertexShader || !fragmentShader) return null;

      const program = gl.createProgram();
      gl.attachShader(program, vertexShader);
      gl.attachShader(program, fragmentShader);
      gl.linkProgram(program);
      gl.deleteShader(vertexShader);
      gl.deleteShader(fragmentShader);

      if (gl.getProgramParameter(program, gl.LINK_STATUS)) return program;
      console.error('[LoadingFluid] program link failed:', gl.getProgramInfoLog(program));
      gl.deleteProgram(program);
      return null;
    };

    /**
     * 绑定一个 float 顶点属性。
     *
     * location < 0 表示该属性被 GLSL 编译器判定为未使用而优化掉了（染料层片元
     * 着色器不吃 uv 时就会这样），此时必须跳过 —— 直接喂 -1 会触发
     * INVALID_VALUE 并在控制台刷 WebGL 报错。
     *
     * @param {number} location 属性位置
     * @param {number} size 分量个数
     * @param {number} stride 步长（字节）
     * @param {number} offset 偏移（字节）
     */
    const bindAttribute = (location, size, stride, offset) => {
      if (location < 0) return;
      gl.enableVertexAttribArray(location);
      gl.vertexAttribPointer(location, size, gl.FLOAT, false, stride, offset);
    };

    /** 建/重建全部 GL 资源。上下文恢复后会被再调一次。 */
    const initResources = () => {
      res.dyeProgram = link(DYE_VERTEX, DYE_FRAGMENT);
      res.planeProgram = link(PLANE_VERTEX, PLANE_FRAGMENT);
      if (!res.dyeProgram || !res.planeProgram) return false;

      res.dyeTexture = gl.createTexture();
      gl.bindTexture(gl.TEXTURE_2D, res.dyeTexture);
      gl.texImage2D(
        gl.TEXTURE_2D, 0, gl.RGBA, RTT_SIZE, RTT_SIZE, 0, gl.RGBA, gl.UNSIGNED_BYTE, null
      );
      // CLAMP_TO_EDGE：波纹把采样坐标推出 [0,1] 时取边缘像素，而不是重复/黑边
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);

      res.dyeFramebuffer = gl.createFramebuffer();
      gl.bindFramebuffer(gl.FRAMEBUFFER, res.dyeFramebuffer);
      gl.framebufferTexture2D(
        gl.FRAMEBUFFER, gl.COLOR_ATTACHMENT0, gl.TEXTURE_2D, res.dyeTexture, 0
      );
      const complete = gl.checkFramebufferStatus(gl.FRAMEBUFFER) === gl.FRAMEBUFFER_COMPLETE;
      gl.bindFramebuffer(gl.FRAMEBUFFER, null);
      if (!complete) return false;

      // 全屏四边形（裁剪空间；uv 仅为与原实现对齐，染料着色器只用 gl_FragCoord）
      res.quadBuffer = gl.createBuffer();
      gl.bindBuffer(gl.ARRAY_BUFFER, res.quadBuffer);
      gl.bufferData(
        gl.ARRAY_BUFFER,
        new Float32Array([-1, -1, 0, 0, 1, -1, 1, 0, -1, 1, 0, 1, 1, 1, 1, 1]),
        gl.STATIC_DRAW
      );

      res.dyeAttribs = {
        position: gl.getAttribLocation(res.dyeProgram, 'position'),
        uv: gl.getAttribLocation(res.dyeProgram, 'uv'),
      };
      res.dyeUniforms = {
        resolution: gl.getUniformLocation(res.dyeProgram, 'u_resolution'),
        time: gl.getUniformLocation(res.dyeProgram, 'u_time'),
        colors: COLORS.map((_, i) => gl.getUniformLocation(res.dyeProgram, 'u_color_' + i)),
        dyes: DYES.map((_, i) => gl.getUniformLocation(res.dyeProgram, 'u_dye_' + i)),
      };
      res.planeAttribs = {
        position: gl.getAttribLocation(res.planeProgram, 'position'),
        uv: gl.getAttribLocation(res.planeProgram, 'uv'),
      };
      res.planeUniforms = {
        map: gl.getUniformLocation(res.planeProgram, 'tMap'),
        time: gl.getUniformLocation(res.planeProgram, 'uTime'),
        magnitude: gl.getUniformLocation(res.planeProgram, 'uMagnitude'),
        modelView: gl.getUniformLocation(res.planeProgram, 'modelViewMatrix'),
        projection: gl.getUniformLocation(res.planeProgram, 'projectionMatrix'),
      };
      return true;
    };

    /** 把染料层画进离屏纹理。与时间无关，只需在（重）建资源后画一次。 */
    const renderDye = () => {
      gl.bindFramebuffer(gl.FRAMEBUFFER, res.dyeFramebuffer);
      gl.viewport(0, 0, RTT_SIZE, RTT_SIZE);
      gl.clearColor(0, 0, 0, 1);
      gl.clear(gl.COLOR_BUFFER_BIT);
      gl.useProgram(res.dyeProgram);

      gl.bindBuffer(gl.ARRAY_BUFFER, res.quadBuffer);
      bindAttribute(res.dyeAttribs.position, 2, 16, 0);
      bindAttribute(res.dyeAttribs.uv, 2, 16, 8);

      gl.uniform2f(res.dyeUniforms.resolution, RTT_SIZE, RTT_SIZE);
      gl.uniform1f(res.dyeUniforms.time, 0);
      COLORS.forEach((hex, i) => gl.uniform3fv(res.dyeUniforms.colors[i], hexToRgb(hex)));
      DYES.forEach((dye, i) => gl.uniform4fv(res.dyeUniforms.dyes[i], dye));

      gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
      gl.bindFramebuffer(gl.FRAMEBUFFER, null);
    };

    if (!initResources()) return null;
    renderDye();

    /* ---------- surface 管理 ---------- */

    /**
     * @type {Map<HTMLCanvasElement, {
     *   ctx: CanvasRenderingContext2D, width: number, height: number,
     *   baseWidth: number, baseHeight: number, baseRatio: number,
     *   projection: Float32Array|null, positions: Float32Array|null,
     *   uvs: Float32Array|null, indices: Uint16Array|null, count: number,
     *   positionBuffer: WebGLBuffer|null, uvBuffer: WebGLBuffer|null,
     *   indexBuffer: WebGLBuffer|null, coverAll: boolean
     * }>}
     */
    const surfaces = new Map();
    let coverAllCount = 0;
    let frame = 0;
    let rafId = 0;
    let running = false;

    const uploadGeometry = (surface) => {
      surface.positionBuffer = gl.createBuffer();
      gl.bindBuffer(gl.ARRAY_BUFFER, surface.positionBuffer);
      gl.bufferData(gl.ARRAY_BUFFER, surface.positions, gl.STATIC_DRAW);

      surface.uvBuffer = gl.createBuffer();
      gl.bindBuffer(gl.ARRAY_BUFFER, surface.uvBuffer);
      gl.bufferData(gl.ARRAY_BUFFER, surface.uvs, gl.STATIC_DRAW);

      surface.indexBuffer = gl.createBuffer();
      gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, surface.indexBuffer);
      gl.bufferData(gl.ELEMENT_ARRAY_BUFFER, surface.indices, gl.STATIC_DRAW);

      surface.count = surface.indices.length;
    };

    /** 把离屏画布撑到能装下所有 surface。只增不减：改尺寸会重新分配后备存储。 */
    const ensureCanvasSize = () => {
      let maxWidth = 0;
      let maxHeight = 0;
      surfaces.forEach((surface) => {
        if (surface.width > maxWidth) maxWidth = surface.width;
        if (surface.height > maxHeight) maxHeight = surface.height;
      });
      if (maxWidth === 0 || maxHeight === 0) return false;
      if (glCanvas.width < maxWidth || glCanvas.height < maxHeight) {
        glCanvas.width = Math.max(glCanvas.width, maxWidth);
        glCanvas.height = Math.max(glCanvas.height, maxHeight);
      }
      return true;
    };

    /** 渲染单个 surface，并把它拷到自己的 2D 画布上。 */
    const drawSurface = (surface) => {
      if (contextLost || !surface.positions || !surface.projection) return;
      if (!ensureCanvasSize()) return;

      gl.bindFramebuffer(gl.FRAMEBUFFER, null);
      // GL 的视口原点在左下角：用 height - h 把 w×h 视口顶到画布**顶部**，
      // 渲染结果就落在「图像左上角」那一块，随后 drawImage(0, 0, w, h) 取到的
      // 正是它，且上下方向已经是正确的（无需翻转）。
      gl.viewport(0, glCanvas.height - surface.height, surface.width, surface.height);

      gl.useProgram(res.planeProgram);
      gl.uniform1i(res.planeUniforms.map, 0);
      gl.uniform1f(res.planeUniforms.magnitude, WAVE_SCALE);
      gl.uniform1f(res.planeUniforms.time, frame / 50);
      gl.uniformMatrix4fv(res.planeUniforms.modelView, false, MODEL_VIEW);
      gl.uniformMatrix4fv(res.planeUniforms.projection, false, surface.projection);

      gl.activeTexture(gl.TEXTURE0);
      gl.bindTexture(gl.TEXTURE_2D, res.dyeTexture);

      gl.bindBuffer(gl.ARRAY_BUFFER, surface.positionBuffer);
      bindAttribute(res.planeAttribs.position, 3, 0, 0);
      gl.bindBuffer(gl.ARRAY_BUFFER, surface.uvBuffer);
      bindAttribute(res.planeAttribs.uv, 2, 0, 0);
      gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, surface.indexBuffer);
      gl.drawElements(gl.TRIANGLES, surface.count, gl.UNSIGNED_SHORT, 0);

      surface.ctx.drawImage(
        glCanvas,
        0, 0, surface.width, surface.height,
        0, 0, surface.width, surface.height
      );
    };

    /** 渲染所有可见 surface。`coverAll` 存在时，被它完全遮住的那些当帧跳过。 */
    const renderAll = () => {
      if (contextLost || !surfaces.size) return;
      surfaces.forEach((surface) => {
        if (coverAllCount > 0 && !surface.coverAll) return;
        drawSurface(surface);
      });
    };

    const removeSurface = (canvas) => {
      const surface = surfaces.get(canvas);
      if (!surface) return;
      surfaces.delete(canvas);
      if (surface.coverAll) coverAllCount -= 1;
      if (observer) observer.unobserve(canvas);
    };

    /** canvas 已离开 DOM（React 接管静态壳、遮罩隐藏……）的 surface 自动摘除。 */
    const pruneDetached = () => {
      surfaces.forEach((_, canvas) => {
        if (!canvas.isConnected) removeSurface(canvas);
      });
    };

    const loop = () => {
      if (!running) return;
      rafId = requestAnimationFrame(loop);
      pruneDetached();
      if (!surfaces.size) {
        stopLoop();
        return;
      }
      frame += 1;
      renderAll();
    };

    const startLoop = () => {
      if (rafId || contextLost || global.document.hidden) return;
      running = true;
      rafId = requestAnimationFrame(loop);
    };

    const stopLoop = () => {
      running = false;
      if (!rafId) return;
      cancelAnimationFrame(rafId);
      rafId = 0;
    };

    /**
     * 同步一个 surface 的尺寸与投影，并立刻出画。
     *
     * 首次拿到有效尺寸时才「真正初始化」它：定下网格边长（max(w, h)）并烘形变。
     * 尺寸还没布局出来（0 / 1px）时直接跳过，等 ResizeObserver 下一次回调。
     *
     * @param {HTMLCanvasElement} canvas surface 画布
     */
    const measure = (canvas) => {
      const surface = surfaces.get(canvas);
      if (!surface) return;
      if (!canvas.isConnected) {
        removeSurface(canvas);
        return;
      }

      const rect = canvas.getBoundingClientRect();
      const width = Math.round(rect.width);
      const height = Math.round(rect.height);
      if (width < 2 || height < 2) return;

      if (width !== surface.width || height !== surface.height) {
        canvas.width = width;
        canvas.height = height;
        surface.width = width;
        surface.height = height;

        if (!surface.positions) {
          surface.baseWidth = width;
          surface.baseHeight = height;
          surface.baseRatio = width / height;
          const geometry = buildPlane(Math.max(width, height));
          surface.positions = geometry.positions;
          surface.uvs = geometry.uvs;
          surface.indices = geometry.indices;
          uploadGeometry(surface);
        }

        // 沿用原库 ColorBg.resize 的 zoom 规则：容器变大时放大视口，画面不跟着抖
        let zoom = 1;
        const ratio = width / height;
        if (ratio > surface.baseRatio) {
          if (width > surface.baseWidth) zoom = width / surface.baseWidth;
        } else if (ratio < surface.baseRatio) {
          if (height > surface.baseHeight) zoom = height / surface.baseHeight;
        }
        surface.projection = orthoMatrix(
          -width / 2 / zoom, width / 2 / zoom,
          -height / 2 / zoom, height / 2 / zoom,
          NEAR, FAR
        );
      }

      // 立刻画一帧：避免「挂载 → 首次 rAF」之间空一帧（静态壳交接时尤其明显）
      drawSurface(surface);
    };

    /* ---------- 尺寸监听 ---------- */

    let observer = null;
    if (typeof ResizeObserver !== 'undefined') {
      observer = new ResizeObserver((entries) => {
        entries.forEach((entry) => measure(entry.target));
      });
    } else {
      // 老浏览器兜底：只能在窗口变化时整批重量
      global.addEventListener('resize', () => {
        surfaces.forEach((_, canvas) => measure(canvas));
      });
    }

    /* ---------- 上下文丢失 / 恢复 ---------- */

    glCanvas.addEventListener('webglcontextlost', (event) => {
      // 必须 preventDefault，否则浏览器不会尝试恢复
      event.preventDefault();
      contextLost = true;
      stopLoop();
    });

    glCanvas.addEventListener('webglcontextrestored', () => {
      if (!initResources()) return;
      contextLost = false;
      renderDye();
      surfaces.forEach((surface) => {
        if (surface.positions) uploadGeometry(surface);
      });
      surfaces.forEach((surface) => drawSurface(surface));
      startLoop();
    });

    global.document.addEventListener('visibilitychange', () => {
      if (global.document.hidden) stopLoop();
      else if (surfaces.size) startLoop();
    });

    /* ---------- 对外接口 ---------- */

    /**
     * 登记一个 surface。
     *
     * @param {HTMLCanvasElement} canvas 展示流体的画布（需已在 DOM 中，且未被
     *   其它上下文占用）
     * @param {{coverAll?: boolean}} [options] coverAll：本层不透明盖满整个可视区
     *   （如启动遮罩），存在时引擎当帧跳过其余被遮挡的 surface
     * @returns {{release: () => void, setCoverAll: (value: boolean) => void}|null}
     */
    const mount = (canvas, options) => {
      const ctx = canvas.getContext('2d');
      if (!ctx) return null;

      const surface = {
        ctx,
        width: 0,
        height: 0,
        baseWidth: 0,
        baseHeight: 0,
        baseRatio: 1,
        projection: null,
        positions: null,
        uvs: null,
        indices: null,
        count: 0,
        positionBuffer: null,
        uvBuffer: null,
        indexBuffer: null,
        coverAll: !!(options && options.coverAll),
      };

      surfaces.set(canvas, surface);
      if (surface.coverAll) coverAllCount += 1;
      if (observer) observer.observe(canvas);

      measure(canvas);
      startLoop();

      let released = false;
      return {
        /**
         * 改「是否盖满整屏」。启动遮罩淡出时必须置回 false ——
         * 那一刻它已经不挡任何东西了，被它盖住的桌面壁纸要立刻开始画，
         * 否则 400ms 的淡出会露出一个空白的桌面，最后才「啪」地补上壁纸。
         *
         * @param {boolean} value
         */
        setCoverAll(value) {
          const next = !!value;
          if (released || next === surface.coverAll) return;
          surface.coverAll = next;
          coverAllCount += next ? 1 : -1;
          // 遮挡解除：下层此前一直没渲染，立刻补一帧
          if (coverAllCount === 0) {
            frame += 1;
            renderAll();
          }
        },

        release() {
          if (released) return;
          released = true;
          const wasCoverAll = surface.coverAll;
          removeSurface(canvas);
          if (wasCoverAll && coverAllCount === 0) {
            frame += 1;
            renderAll();
          }
          if (!surfaces.size) stopLoop();
        },
      };
    };

    return { mount };
  }

  /**
   * 取全局单例（懒建）。失败会被记成 false，后续调用不再重试。
   *
   * @returns {object|null}
   */
  function getEngine() {
    if (engine === null) engine = createEngine() || false;
    return engine || null;
  }

  global.LoadingFluid = {
    /**
     * 引擎是否可用（会触发一次懒建）。
     *
     * @returns {boolean}
     */
    supported() {
      return !!getEngine();
    },

    /**
     * 挂载一个流体 surface。
     *
     * @param {HTMLCanvasElement} canvas 已插入 DOM 的展示用画布
     * @param {{coverAll?: boolean}} [options]
     * @returns {{release: () => void, setCoverAll: (value: boolean) => void}|null}
     *   引擎不可用时返回 null（调用方自行降级）
     */
    mount(canvas, options) {
      const instance = getEngine();
      if (!instance || !canvas) return null;
      return instance.mount(canvas, options || {});
    },
  };
})(window);
