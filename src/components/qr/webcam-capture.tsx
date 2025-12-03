import React, { useEffect, useRef, useState } from "react";
import Webcam from "react-webcam";

/**
 * Draws from the live video element to a hidden canvas to produce ImageData
 * and calls onScan(ImageData). Keeps resolution high for better decoding.
 */
export default function WebcamCapture({ onScan, interval = 400 }) {
  const webcamRef = useRef(null);
  const canvasRef = useRef(null);
  const [stream, setStream] = useState(null);
  const [torchAvailable, setTorchAvailable] = useState(false);
  const [torchOn, setTorchOn] = useState(false);

  useEffect(() => {
    const id = setInterval(() => {
      capture();
    }, interval);
    return () => clearInterval(id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [webcamRef.current, interval, stream, torchOn]);

  useEffect(() => {
    const v = webcamRef.current?.video;
    const s = v?.srcObject;
    if (!s) return;
    const t = s.getVideoTracks?.()?.[0];
    if (!t) { setTorchAvailable(false); return; }
    try {
      const caps = t.getCapabilities?.();
      setTorchAvailable(Boolean(caps && (caps.torch || caps.torch === true)));
    } catch (e) {
      setTorchAvailable(false);
    }
  }, [stream]);

  const capture = () => {
    const video = webcamRef.current?.video;
    if (!video || video.readyState < 2) return;

    let canvas = canvasRef.current;
    if (!canvas) {
      canvas = document.createElement("canvas");
      canvasRef.current = canvas;
    }

    const vw = video.videoWidth || 1280;
    const vh = video.videoHeight || 720;

    // keep a reasonable maximum size so mobile CPU doesn't overwork
    const targetWidth = Math.min(vw, 1600);
    const aspect = vh / vw;
    const targetHeight = Math.round(targetWidth * aspect);

    canvas.width = targetWidth;
    canvas.height = targetHeight;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });

    try {
      ctx.drawImage(video, 0, 0, targetWidth, targetHeight);
    } catch (e) {
      return;
    }

    let imageData;
    try {
      imageData = ctx.getImageData(0, 0, targetWidth, targetHeight);
    } catch (e) {
      // if canvas is tainted for any reason, skip
      return;
    }

    onScan?.(imageData);
  };

  const toggleTorch = async () => {
    const video = webcamRef.current?.video;
    const s = video?.srcObject;
    const track = s?.getVideoTracks?.()?.[0];
    if (!track || !("applyConstraints" in track)) return;
    try {
      await track.applyConstraints({ advanced: [{ torch: !torchOn }] });
      setTorchOn((v) => !v);
    } catch (e) {
      // ignore
    }
  };

  const videoConstraints = {
    width: { ideal: 1920 },
    height: { ideal: 1080 },
    facingMode: "environment",
  };

  return (
    <div>
      <Webcam
        ref={webcamRef}
        audio={false}
        screenshotFormat="image/png"
        videoConstraints={videoConstraints}
        onUserMedia={(s) => setStream(s)}
        mirrored={false}
        style={{ width: "100%", height: "auto", maxHeight: 480, background: "#000" }}
      />
      <div className="mt-2 flex gap-3">
        {torchAvailable && (
          <button onClick={toggleTorch} className="flex-1 py-3 rounded-lg bg-yellow-500 text-black">
            {torchOn ? "Torch off" : "Torch on"}
          </button>
        )}
        <button onClick={capture} className="flex-1 py-3 rounded-lg bg-blue-600 text-white">
          Capture
        </button>
      </div>
    </div>
  );
}