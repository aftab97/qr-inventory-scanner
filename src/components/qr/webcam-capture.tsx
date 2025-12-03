import { useEffect, useRef, useState } from "react";
import Webcam from "react-webcam";

type WebcamCaptureProps = {
  onScan?: (imageData: ImageData) => void;
  interval?: number;
};

/**
 * Minimal typed subset of the react-webcam instance we use.
 * react-webcam's instance exposes a `.video` element and a `getScreenshot` method.
 * We only need `.video` here.
 */
type WebcamInstanceLike = {
  video?: HTMLVideoElement | null;
  getScreenshot?: (params?: { width?: number; height?: number }) => string | null;
};

export default function WebcamCapture({ onScan, interval = 400 }: WebcamCaptureProps) {
  // we keep a narrow, explicit type for the component instance
  const webcamRef = useRef<WebcamInstanceLike | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [stream, setStream] = useState<MediaStream | null>(null);
  const [torchAvailable, setTorchAvailable] = useState<boolean>(false);
  const [torchOn, setTorchOn] = useState<boolean>(false);

  useEffect(() => {
    const id = window.setInterval(() => {
      capture();
    }, interval);
    return () => window.clearInterval(id);
    // intentionally not depending on capture reference
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [interval, stream, torchOn]);

  useEffect(() => {
    const v = webcamRef.current?.video;
    const s = v?.srcObject as MediaStream | undefined | null;
    if (!s) return;
    const track = s.getVideoTracks?.()?.[0];
    if (!track) {
      setTorchAvailable(false);
      return;
    }
    try {
      const caps = (track as MediaStreamTrack & { getCapabilities?: () => MediaTrackCapabilities }).getCapabilities?.();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      setTorchAvailable(Boolean(caps && (caps as any).torch)); // capability shape is not uniform across browsers
    } catch {
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
    const aspect = vh / vw || 9 / 16;
    const targetHeight = Math.max(1, Math.round(targetWidth * aspect));

    canvas.width = targetWidth;
    canvas.height = targetHeight;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    if (!ctx) return;

    try {
      ctx.drawImage(video, 0, 0, targetWidth, targetHeight);
    } catch {
      // drawing may fail if video is not ready or something else went wrong
      return;
    }

    let imageData: ImageData;
    try {
      imageData = ctx.getImageData(0, 0, targetWidth, targetHeight);
    } catch {
      // if canvas is tainted for any reason, skip
      return;
    }

    if (onScan) onScan(imageData);
  };

  const toggleTorch = async () => {
    const video = webcamRef.current?.video;
    const s = video?.srcObject as MediaStream | undefined | null;
    const track = s?.getVideoTracks?.()?.[0];
    if (!track || typeof track.applyConstraints !== "function") return;
    try {
      // applyConstraints may fail on browsers that don't support torch
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore - browser-specific constraint shape
      await track.applyConstraints({ advanced: [{ torch: !torchOn }] });
      setTorchOn((v) => !v);
    } catch {
      // ignore failures toggling torch
    }
  };

  const handleUserMedia = (mediaStream: MediaStream) => {
    setStream(mediaStream);
    // also attempt to populate the webcamRef.video shortly after media stream is available
    setTimeout(() => {
      // no-op; the webcamRef will be populated by the ref callback below
    }, 200);
  };

  const videoConstraints: MediaTrackConstraints = {
    width: { ideal: 1920 },
    height: { ideal: 1080 },
    facingMode: "environment",
  };

  return (
    <div>
      <Webcam
        // react-webcam's ref is an instance; we assign it to our typed ref via a callback
        ref={(instance) => {
          // instance is of unknown type in the library; coerce into our narrow interface
          webcamRef.current = instance as unknown as WebcamInstanceLike | null;
        }}
        audio={false}
        screenshotFormat="image/png"
        videoConstraints={videoConstraints}
        onUserMedia={handleUserMedia}
        mirrored={false}
        style={{ width: "100%", height: "auto", maxHeight: 480, background: "#000" }}
      />
      <div className="mt-2 flex gap-3">
        {torchAvailable && (
          <button onClick={toggleTorch} className="flex-1 py-3 rounded-lg bg-yellow-500 text-black" type="button">
            {torchOn ? "Torch off" : "Torch on"}
          </button>
        )}
        <button onClick={capture} className="flex-1 py-3 rounded-lg bg-blue-600 text-white" type="button">
          Capture
        </button>
      </div>
    </div>
  );
}