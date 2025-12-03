// See earlier message for comments — small, fast helpers used server-side-ish in browser.

export function enhanceContrast(imageData, contrast = 30) {
  const factor = (259 * (contrast + 255)) / (255 * (259 - contrast));
  const data = imageData.data;
  for (let i = 0; i < data.length; i += 4) {
    data[i] = clamp(Math.round(factor * (data[i] - 128) + 128));
    data[i + 1] = clamp(Math.round(factor * (data[i + 1] - 128) + 128));
    data[i + 2] = clamp(Math.round(factor * (data[i + 2] - 128) + 128));
  }
  return imageData;
}

export function unsharpMask(imageData, amount = 1.0) {
  const w = imageData.width;
  const h = imageData.height;
  const src = new Uint8ClampedArray(imageData.data);
  const dst = imageData.data;
  const k = [
    [0, -1, 0],
    [-1, 5, -1],
    [0, -1, 0],
  ];
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      let r = 0, g = 0, b = 0;
      for (let ky = -1; ky <= 1; ky++) {
        for (let kx = -1; kx <= 1; kx++) {
          const sx = clampX(x + kx, w);
          const sy = clampY(y + ky, h);
          const idx = (sy * w + sx) * 4;
          const kval = k[ky + 1][kx + 1] * amount;
          r += src[idx] * kval;
          g += src[idx + 1] * kval;
          b += src[idx + 2] * kval;
        }
      }
      const idx = (y * w + x) * 4;
      dst[idx] = clamp(Math.round(r));
      dst[idx + 1] = clamp(Math.round(g));
      dst[idx + 2] = clamp(Math.round(b));
    }
  }
  return imageData;
}

export function invertImageData(imageData) {
  const w = imageData.width;
  const h = imageData.height;
  const src = imageData.data;
  const out = new ImageData(w, h);
  for (let i = 0; i < src.length; i += 4) {
    out.data[i] = 255 - src[i];
    out.data[i + 1] = 255 - src[i + 1];
    out.data[i + 2] = 255 - src[i + 2];
    out.data[i + 3] = src[i + 3] ?? 255;
  }
  return out;
}

export function adaptiveThresholdBradley(srcImage, windowSize = 41, t = 0.15) {
  const w = srcImage.width, h = srcImage.height;
  const data = srcImage.data;
  const gray = new Uint32Array(w * h);
  for (let i = 0, p = 0; i < data.length; i += 4, p++) {
    gray[p] = Math.round(0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]);
  }

  const integral = new Uint32Array((w + 1) * (h + 1));
  for (let y = 1; y <= h; y++) {
    let rowSum = 0;
    for (let x = 1; x <= w; x++) {
      rowSum += gray[(y - 1) * w + (x - 1)];
      integral[y * (w + 1) + x] = integral[(y - 1) * (w + 1) + x] + rowSum;
    }
  }

  const out = new ImageData(w, h);
  const half = Math.floor(windowSize / 2);
  for (let y = 0; y < h; y++) {
    const y0 = Math.max(0, y - half);
    const y1 = Math.min(h - 1, y + half);
    for (let x = 0; x < w; x++) {
      const x0 = Math.max(0, x - half);
      const x1 = Math.min(w - 1, x + half);
      const count = (x1 - x0 + 1) * (y1 - y0 + 1);
      const sum =
        integral[(y1 + 1) * (w + 1) + (x1 + 1)] -
        integral[(y0) * (w + 1) + (x1 + 1)] -
        integral[(y1 + 1) * (w + 1) + (x0)] +
        integral[(y0) * (w + 1) + x0];
      const mean = sum / count;
      const idx = y * w + x;
      const val = gray[idx];
      const outv = val < mean * (1 - t) ? 0 : 255;
      const i4 = idx * 4;
      out.data[i4] = out.data[i4 + 1] = out.data[i4 + 2] = outv;
      out.data[i4 + 3] = 255;
    }
  }
  return out;
}

export function binaryClose(imageData, iterations = 1) {
  const w = imageData.width, h = imageData.height;
  const src = new Uint8ClampedArray((imageData.data.length / 4));
  for (let i = 0, p = 0; i < imageData.data.length; i += 4, p++) {
    src[p] = imageData.data[i] > 128 ? 1 : 0;
  }

  const tmp = new Uint8ClampedArray(src.length);

  const dilate = (inArr, outArr) => {
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        let on = 0;
        for (let ky = -1; ky <= 1 && !on; ky++) {
          for (let kx = -1; kx <= 1; kx++) {
            const sx = x + kx;
            const sy = y + ky;
            if (sx < 0 || sx >= w || sy < 0 || sy >= h) continue;
            if (inArr[sy * w + sx]) { on = 1; break; }
          }
        }
        outArr[y * w + x] = on;
      }
    }
  };

  const erode = (inArr, outArr) => {
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        let all = 1;
        for (let ky = -1; ky <= 1 && all; ky++) {
          for (let kx = -1; kx <= 1; kx++) {
            const sx = x + kx;
            const sy = y + ky;
            if (sx < 0 || sx >= w || sy < 0 || sy >= h) { all = 0; break; }
            if (!inArr[sy * w + sx]) { all = 0; break; }
          }
        }
        outArr[y * w + x] = all;
      }
    }
  };

  let cur = src, work = tmp;
  for (let it = 0; it < iterations; it++) { dilate(cur, work); const t = cur; cur = work; work = t; }
  for (let it = 0; it < iterations; it++) { erode(cur, work); const t = cur; cur = work; work = t; }

  const out = new ImageData(w, h);
  for (let p = 0; p < cur.length; p++) {
    const v = cur[p] ? 255 : 0;
    const i4 = p * 4;
    out.data[i4] = out.data[i4 + 1] = out.data[i4 + 2] = v;
    out.data[i4 + 3] = 255;
  }
  return out;
}

/* helpers */
function clamp(v) { return v < 0 ? 0 : v > 255 ? 255 : v | 0; }
function clampX(x, w) { if (x < 0) return 0; if (x >= w) return w - 1; return x; }
function clampY(y, h) { if (y < 0) return 0; if (y >= h) return h - 1; return y; }