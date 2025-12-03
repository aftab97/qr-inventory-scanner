import React, { useCallback, useState } from "react";
import jsQR from "jsqr";
import WebcamCapture from "./webcam-capture";
import {
  enhanceContrast,
  unsharpMask,
  adaptiveThresholdBradley,
  invertImageData,
  binaryClose,
} from "./image-processing";

/**
 * QRScanner component
 * - receives ImageData from WebcamCapture
 * - runs multiple preprocessing strategies (fast -> heavier)
 * - tries jsQR for each strategy and reports first success via onResult prop
 */
export default function QRScanner({ onResult }) {
  const [lastStrategy, setLastStrategy] = useState("");
  const [decoded, setDecoded] = useState(null);

  const handleScan = useCallback(
    (imageData) => {
      if (!imageData) return;

      const strategies = [
        { name: "raw", fn: (id) => id },
        {
          name: "enhanced",
          fn: (id) => {
            const c = new ImageData(new Uint8ClampedArray(id.data), id.width, id.height);
            enhanceContrast(c, 36);
            unsharpMask(c, 1.0);
            return c;
          },
        },
        {
          name: "adaptive",
          fn: (id) => {
            const c = new ImageData(new Uint8ClampedArray(id.data), id.width, id.height);
            enhanceContrast(c, 36);
            unsharpMask(c, 1.0);
            return adaptiveThresholdBradley(c, 41, 0.12);
          },
        },
        {
          name: "adaptive-close",
          fn: (id) => binaryClose(adaptiveThresholdBradley(new ImageData(new Uint8ClampedArray(id.data), id.width, id.height), 41, 0.12), 1),
        },
        {
          name: "invert-adaptive",
          fn: (id) => adaptiveThresholdBradley(invertImageData(new ImageData(new Uint8ClampedArray(id.data), id.width, id.height)), 41, 0.12),
        },
      ];

      for (const strat of strategies) {
        try {
          const processed = strat.fn(imageData);
          const code = jsQR(processed.data, processed.width, processed.height, { inversionAttempts: "dontInvert" });
          if (code) {
            setDecoded(code.data ?? code);
            setLastStrategy(strat.name);
            onResult?.(code.data ?? code);
            return;
          }
        } catch (e) {
          // strategy failed, try next
          // eslint-disable-next-line no-console
          console.warn("Strategy failed", strat.name, e);
        }
      }

      // fallback to jsQR attempts
      try {
        const fb = jsQR(imageData.data, imageData.width, imageData.height, { inversionAttempts: "attemptBoth" });
        if (fb) {
          setDecoded(fb.data ?? fb);
          setLastStrategy("fallback-jsqr");
          onResult?.(fb.data ?? fb);
        }
      } catch (e) {
        // ignore
      }
    },
    [onResult]
  );

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between mb-2">
        <div>
          <div className="text-sm text-gray-500">Scanner</div>
          <div className="text-xs text-gray-400">Strategy: {lastStrategy || "—"}</div>
        </div>
        <div className="text-sm text-gray-700 font-medium">{decoded ?? "No result"}</div>
      </div>

      <div className="rounded-lg overflow-hidden">
        <WebcamCapture onScan={handleScan} interval={450} />
      </div>
    </div>
  );
}