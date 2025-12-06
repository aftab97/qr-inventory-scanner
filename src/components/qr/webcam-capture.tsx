import React, { useEffect, useRef, useState, forwardRef, useImperativeHandle } from "react";
import Webcam from "react-webcam";

type WebcamCaptureProps = {
  onScan?: (imageData: ImageData) => void;
  interval?: number; // ms
};

export type WebcamCaptureHandle = {
  stop: () => void;
};

type WebcamInstanceLike = {
  video?: HTMLVideoElement | null;
};

const WebcamCapture = forwardRef<WebcamCaptureHandle, WebcamCaptureProps>(function WebcamCapture(
  { onScan, interval = 400 },
  ref
) {
  const webcamRef = useRef<WebcamInstanceLike | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [stream, setStream] = useState<MediaStream | null>(null);
  const [timerId, setTimerId] = useState<number | null>(null);
  const [torchAvailable, setTorchAvailable] = useState<boolean>(false);
  const [torchOn, setTorchOn] = useState<boolean>(false);

  useImperativeHandle(ref, () => ({
    stop: () => {
      // stop interval
      if (timerId) {
        window.clearInterval(timerId);
        setTimerId(null);
      }
      // stop media tracks
      try {
        const v = webcamRef.current?.video;
        const s = v?.srcObject as MediaStream | undefined | null;
        s?.getTracks()?.forEach((t) => t.stop());
      } catch {
        // ignore
      }
    },
  }));

  useEffect(() => {
    // start periodic capture when stream becomes available
    if (!stream || timerId) return;
    const id = window.setInterval(() => capture(), interval);
    setTimerId(id as unknown as number);
    return () => {
      window.clearInterval(id);
      setTimerId(null);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [stream, interval]);

  // detect torch support
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
      setTorchAvailable(Boolean(caps && (caps as any).torch));
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
      return;
    }

    let imageData: ImageData;
    try {
      imageData = ctx.getImageData(0, 0, targetWidth, targetHeight);
    } catch {
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
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      // @ts-ignore
      await track.applyConstraints({ advanced: [{ torch: !torchOn }] });
      setTorchOn((v) => !v);
    } catch {
      // ignore
    }
  };

  const handleUserMedia = (mediaStream: MediaStream) => {
    setStream(mediaStream);
  };

  const videoConstraints: MediaTrackConstraints = {
    width: { ideal: 1920 },
    height: { ideal: 1080 },
    facingMode: "environment",
  };

  return (
    <div>
      <Webcam
        ref={(instance) => {
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
        {/* No manual capture button — auto-capturing via interval */}
      </div>
    </div>
  );
});

export default WebcamCapture;