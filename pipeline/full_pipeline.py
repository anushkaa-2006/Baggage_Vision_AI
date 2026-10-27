import cv2
import numpy as np
import tempfile
from pathlib import Path

from inference.detect_objects import ObjectDetector
from inference.risk_mapping import compute_risk, get_decision
from preprocessing.image_preprocessing import preprocess_image

from frontend.anomaly_detection.feature_extractor import extract_feature
from frontend.anomaly_detection.anomaly_detector import AnomalyDetector

PROJECT_ROOT = Path(__file__).resolve().parent.parent


class BaggageScreeningPipeline:

    def __init__(self):
        self.detector = ObjectDetector()

        feature_path = (
            PROJECT_ROOT
            / "frontend"
            / "anomaly_detection"
            / "normal_features.npy"
        )

        self.anomaly_detector = AnomalyDetector(str(feature_path))

    def preprocess_and_save(self, image_path):
        processed = preprocess_image(image_path)

        with tempfile.NamedTemporaryFile(
            suffix=".jpg", delete=False
        ) as temp_file:
            temp_path = temp_file.name

        if not cv2.imwrite(temp_path, processed):
            Path(temp_path).unlink(missing_ok=True)
            raise RuntimeError("Could not save the preprocessed X-ray image.")

        return temp_path

    def anomaly_score(self, image_path):
        feature = extract_feature(image_path)
        distance = self.anomaly_detector.compute_score(feature)

        # Larger distance from normal training features means
        # a higher anomaly score.
        score = 1 - np.exp(-float(distance) / 5)
        return float(np.clip(score, 0.0, 1.0))

    def run(self, image_path):
        processed_path = self.preprocess_and_save(image_path)

        try:
            detections, detection_score = self.detector.detect(
                processed_path
            )

            anomaly = self.anomaly_score(processed_path)

            # No cargo manifest is used in baggage screening.
            mismatch = 0.0

            risk = compute_risk(
                detection_score,
                anomaly,
                mismatch
            )

            decision = get_decision(risk)

            return {
                "detections": detections,
                "scores": {
                    "detection": float(detection_score),
                    "anomaly": float(anomaly),
                    "mismatch": float(mismatch),
                    "risk": float(risk),
                },
                "decision": decision,
                "processed_image_path": processed_path,
            }

        except Exception:
            Path(processed_path).unlink(missing_ok=True)
            raise
