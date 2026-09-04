/**
 * ClickSpark - React Bits Vanilla JS Port
 * Interactive click spark burst effect generating radiant spark lines on user click.
 */

function initClickSpark(options = {}) {
  if (window.__clickSparkInitialized) return;
  window.__clickSparkInitialized = true;

  const {
    sparkColor = null,
    sparkSize = 10,
    sparkRadius = 15,
    sparkCount = 8,
    duration = 400,
    easing = 'ease-out',
    extraScale = 1.0
  } = options;

  const canvas = document.createElement('canvas');
  canvas.className = 'click-spark-canvas';
  canvas.style.position = 'fixed';
  canvas.style.top = '0';
  canvas.style.left = '0';
  canvas.style.width = '100vw';
  canvas.style.height = '100vh';
  canvas.style.pointerEvents = 'none';
  canvas.style.zIndex = '999999';
  canvas.style.userSelect = 'none';
  canvas.style.display = 'block';
  document.body.appendChild(canvas);

  const ctx = canvas.getContext('2d');
  let sparks = [];

  const resize = () => {
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    canvas.width = window.innerWidth * dpr;
    canvas.height = window.innerHeight * dpr;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  };
  window.addEventListener('resize', resize);
  resize();

  const easeFunc = t => {
    switch (easing) {
      case 'linear':
        return t;
      case 'ease-in':
        return t * t;
      case 'ease-in-out':
        return t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t;
      default:
        return t * (2 - t);
    }
  };

  let animationId = null;

  const draw = timestamp => {
    ctx.clearRect(0, 0, window.innerWidth, window.innerHeight);

    const isDark = document.documentElement.classList.contains('dark');
    const activeColor = sparkColor || (isDark ? '#2dd4bf' : '#247568');

    sparks = sparks.filter(spark => {
      const elapsed = timestamp - spark.startTime;
      if (elapsed >= duration) {
        return false;
      }

      const progress = elapsed / duration;
      const eased = easeFunc(progress);

      const distance = eased * sparkRadius * extraScale;
      const lineLength = sparkSize * (1 - eased);

      const x1 = spark.x + distance * Math.cos(spark.angle);
      const y1 = spark.y + distance * Math.sin(spark.angle);
      const x2 = spark.x + (distance + lineLength) * Math.cos(spark.angle);
      const y2 = spark.y + (distance + lineLength) * Math.sin(spark.angle);

      ctx.strokeStyle = activeColor;
      ctx.lineWidth = 2;
      ctx.lineCap = 'round';
      ctx.beginPath();
      ctx.moveTo(x1, y1);
      ctx.lineTo(x2, y2);
      ctx.stroke();

      return true;
    });

    if (sparks.length > 0) {
      animationId = requestAnimationFrame(draw);
    } else {
      ctx.clearRect(0, 0, window.innerWidth, window.innerHeight);
      animationId = null;
    }
  };

  window.addEventListener('pointerdown', e => {
    const x = e.clientX;
    const y = e.clientY;
    const now = performance.now();

    for (let i = 0; i < sparkCount; i++) {
      sparks.push({
        x,
        y,
        angle: (2 * Math.PI * i) / sparkCount,
        startTime: now
      });
    }

    if (!animationId) {
      animationId = requestAnimationFrame(draw);
    }
  }, { passive: true });
}

window.initClickSpark = initClickSpark;

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => initClickSpark());
} else {
  initClickSpark();
}
