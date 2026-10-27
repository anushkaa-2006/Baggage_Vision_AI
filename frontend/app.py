
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.append(
    str(Path(__file__).resolve().parent.parent)
)

from pipeline.full_pipeline import BaggageScreeningPipeline
from utils.visualization import draw_boxes


st.set_page_config(
    page_title="Baggage Vision AI",
    page_icon="🔎",
    layout="wide",
)

st.markdown("""
<style>
.block-container {
    padding-top: 2rem;
    max-width: 1400px;
}

.title {
    font-size: 38px;
    font-weight: 800;
    color: #2563eb;
}

.subtitle {
    color: #9ca3af;
    font-size: 16px;
    margin-bottom: 24px;
}

.panel {
    background: #1e293b;
    padding: 20px;
    border-radius: 14px;
    margin-bottom: 18px;
}

.metric {
    background: #1e293b;
    padding: 18px;
    border-radius: 12px;
    text-align: center;
    border: 1px solid #334155;
}

.safe {
    background: #166534;11
    color: white;
    padding: 18px;
    border-radius: 10px;
    font-size: 22px;
    font-weight: 700;
    text-align: center;
}

.warn {
    background: #a16207;
    color: white;
    padding: 18px;
    border-radius: 10px;
    font-size: 22px;
    font-weight: 700;
    text-align: center;
}

.danger {
    background: #991b1b;
    color: white;
    padding: 18px;
    border-radius: 10px;
    font-size: 22px;
    font-weight: 700;
    text-align: center;
}
</style>
""", unsafe_allow_html=True)


st.markdown(
    '<div class="title">Baggage Vision AI</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitle">'
    'Intelligent Luggage Inspection and Risk Detection System. | '
    'Airport • Metro • Railway'
    '</div>',
    unsafe_allow_html=True,
)

st.info(
    "Upload one baggage X-ray image to run AI screening. "
    "Results are decision-support outputs and require human "
    "review; they are not a definitive security clearance."
)

# Cache the model instance so it is not reloaded on every rerun.
@st.cache_resource
def load_pipeline():
    return BaggageScreeningPipeline()


try:
    pipeline = load_pipeline()
except Exception as exc:
    st.error(f"Could not initialize the screening models: {exc}")
    st.stop()


st.markdown('<div class="panel">', unsafe_allow_html=True)

uploaded_image = st.file_uploader(
    "Upload baggage X-ray image",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=False,
    help="Upload one X-ray image in PNG or JPEG format.",
)

if uploaded_image is not None:
    _, _, preview_col, _, _ = st.columns([1, 1, 2, 1, 1])

    with preview_col:
        st.image(
            uploaded_image.getvalue(),
            caption="Uploaded baggage X-ray",
            width=300,
        )

analyze = st.button(
    "Analyze X-ray",
    type="primary",
    use_container_width=True,
    disabled=uploaded_image is None,
)

st.markdown('</div>', unsafe_allow_html=True)


if analyze and uploaded_image is not None:
    input_path = None
    processed_path = None

    try:
        suffix = Path(uploaded_image.name).suffix.lower()
        if suffix not in {".png", ".jpg", ".jpeg"}:
            suffix = ".jpg"

        with tempfile.NamedTemporaryFile(
            suffix=suffix, delete=False
        ) as temp_file:
            temp_file.write(uploaded_image.getvalue())
            input_path = temp_file.name

        with st.spinner("Analyzing baggage X-ray..."):
            result = pipeline.run(input_path)

            detections = result["detections"]
            scores = result["scores"]
            decision = result["decision"]
            processed_path = result.get("processed_image_path")

            annotated = draw_boxes(
                processed_path or input_path,
                detections,
            )

        
        st.divider()
        st.subheader("Screening Results")

        # Three compact columns across the page
        image_col, objects_col, risk_col = st.columns(
            [1.15, 1, 0.85],
            gap="medium"
        )

        with image_col:
            st.markdown("**X-ray Detection**")
            st.image(
                annotated,
                caption="AI-detected objects",
                width=350,
            )

        with objects_col:
            st.markdown("**Detected Objects**")

            if detections:
                df = pd.DataFrame([
                    {
                        "Object": d.get(
                            "class", d.get("object", "Unknown")
                        ),
                        "Confidence (%)": round(
                            float(d.get("confidence", 0)) * 100, 1
                        ),
                    }
                    for d in detections
                ])

                st.dataframe(
                    df,
                    use_container_width=True,
                    hide_index=True,
                    height=220,
                )
            else:
                st.info("No objects detected.")

            st.markdown("**AI Scores**")

            st.metric(
                "Object Detection",
                f'{scores["detection"]:.2f}'
            )
            st.metric(
                "Anomaly Score",
                f'{scores["anomaly"]:.2f}'
            )
            st.metric(
                "Overall Risk",
                f'{scores["risk"]:.2f}'
            )

        with risk_col:
            st.markdown("**Screening Decision**")

            if decision == "SAFE":
                st.success("LOW RISK")
            elif decision == "SUSPICIOUS":
                st.warning("REVIEW REQUIRED")
            else:
                st.error("HIGH RISK")

            st.caption(f"Pipeline decision: {decision}")

            st.markdown("---")
            st.markdown("**Model Contributions**")

            st.write("Object Detection")
            st.progress(
                max(0.0, min(1.0, float(scores["detection"])))
            )

            st.write("Anomaly Detection")
            st.progress(
                max(0.0, min(1.0, float(scores["anomaly"])))
            )

            st.caption(
                "AI results support human review; they do not "
                "provide definitive security clearance."
            )


    except Exception as exc:
        st.error(f"X-ray analysis failed: {exc}")

    finally:
        for temp_path in (input_path, processed_path):
            if temp_path:
                try:
                    os.unlink(temp_path)
                except OSError:
                    pass
