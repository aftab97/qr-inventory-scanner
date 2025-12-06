import React, { useCallback, useRef, useState, forwardRef, useImperativeHandle } from "react";
import jsQR from "jsqr";
import {
  enhanceContrast,
  unsharpMask,
  adaptiveThresholdBradley,
  invertImageData,
  binaryClose,
} from "./image-processing";
import WebcamCapture, { type WebcamCaptureHandle } from "./webcam-capture";

type JsQRResult = { data: string; location?: unknown } | null;

type QRScannerProps = {
  onResult?: (text: string) => void;
  onStop?: () => void; // notify parent when scanning stops
};

export type QRScannerHandle = {
  stop: () => void;
};

const QRScanner = forwardRef<QRScannerHandle, QRScannerProps>(function QRScanner({ onResult, onStop }, ref) {
  const [lastStrategy, setLastStrategy] = useState<string>("");
  const [decoded, setDecoded] = useState<string | null>(null);
  const webcamRef = useRef<WebcamCaptureHandle | null>(null);

  // jsQR typed facade
  const jsqrDecode = (jsQR as unknown) as (
    data: Uint8ClampedArray,
    width: number,
    height: number,
    opts?: { inversionAttempts?: "dontInvert" | "attemptBoth" | string }
  ) => JsQRResult;

  useImperativeHandle(ref, () => ({
    stop: () => {
      try {
        webcamRef.current?.stop();
        onStop?.();
      } catch {
        // ignore
      }
    },
  }));

  const stopScanning = useCallback(() => {
    try {
      webcamRef.current?.stop();
    } catch {}
    onStop?.();
  }, [onStop]);

  const handleScan = useCallback(
    (imageData: ImageData | null) => {
      if (!imageData) return;

      const strategies: { name: string; fn: (id: ImageData) => ImageData }[] = [
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
          fn: (id) =>
            binaryClose(adaptiveThresholdBradley(new ImageData(new Uint8ClampedArray(id.data), id.width, id.height), 41, 0.12), 1),
        },
        {
          name: "invert-adaptive",
          fn: (id) => adaptiveThresholdBradley(invertImageData(new ImageData(new Uint8ClampedArray(id.data), id.width, id.height)), 41, 0.12),
        },
      ];

      // Try strategies; if any decodes, stop scanning immediately
      for (const strat of strategies) {
        try {
          const processed = strat.fn(imageData);
          const code = jsqrDecode(processed.data, processed.width, processed.height, { inversionAttempts: "dontInvert" });
          if (code && code.data) {
            setDecoded(code.data);
            setLastStrategy(strat.name);
            onResult?.(code.data);
            // IMPORTANT: stop the webcam immediately after a decode
            stopScanning();
            return;
          }
        } catch (e) {
          // eslint-disable-next-line no-console
          console.warn("Strategy failed", strat.name, e as unknown);
        }
      }

      // fallback attemptBoth
      try {
        const fb = jsqrDecode(imageData.data, imageData.width, imageData.height, { inversionAttempts: "attemptBoth" });
        if (fb && fb.data) {
          setDecoded(fb.data);
          setLastStrategy("fallback-jsqr");
          onResult?.(fb.data);
          stopScanning();
        }
      } catch {
        // ignore
      }
    },
    [onResult, jsqrDecode, stopScanning]
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
        <WebcamCapture ref={webcamRef} onScan={handleScan} interval={450} />
      </div>
    </div>
  );
});

export default QRScanner;