import sys
import os
import streamlit as st
import tempfile
import pandas as pd

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from pipeline.full_pipeline import CargoPipeline
from utils.visualization import draw_boxes
from database import init_db, insert_container, get_latest_container

# ---------------- INIT ----------------
init_db()

st.set_page_config(
    page_title="Cargo Risk Intelligence",
    layout="wide"
)
6
# ---------------- STYLING ----------------
st.markdown("""
<style>
.block-container { padding-top:2rem; }

.main { background-color:#f4f8ff; }

.title {
    font-size:42px;
    font-weight:800;
    color:#0a3d91;
}

.card {
    background:white;
    padding:20px;
    border-radius:12px;
    box-shadow:0px 4px 12px rgba(0,0,0,0.08);
    margin-bottom:20px;
}

.metric-card {
    background:#0a3d91;
    color:white;
    padding:16px;
    border-radius:10px;
    text-align:center;
}

.case-strip {
    background:white;
    padding:10px;
    border-radius:10px;
    margin-bottom:15px;
}

.safe { background:#1abc9c; color:white; padding:15px; border-radius:10px; }
.warn { background:#f39c12; color:white; padding:15px; border-radius:10px; }
.danger { background:#e74c3c; color:white; padding:15px; border-radius:10px; }
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------
st.markdown('<div class="title">AI Cargo Risk Intelligence System</div>', unsafe_allow_html=True)

# ---------------- CASE STRIP ----------------
latest = get_latest_container()

if latest:
    st.markdown(f"""
    <div class="case-strip">
    <b>{latest[1]}</b> | {latest[2]} → {latest[3]} | Container: {latest[4]}
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown('<div class="case-strip">No active container</div>', unsafe_allow_html=True)

pipeline = CargoPipeline()

# ---------------- INPUT ----------------
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="card">', unsafe_allow_html=True)

    case_id = st.text_input("Case ID")
    origin = st.text_input("Origin")
    destination = st.text_input("Destination")
    container_no = st.text_input("Container No")

    if st.button("Save Container"):
        insert_container(case_id, origin, destination, container_no)
        st.success("Saved")

    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="card">', unsafe_allow_html=True)

    uploaded_image = st.file_uploader("Upload X-ray", type=["png","jpg","jpeg"])
    uploaded_manifest = st.file_uploader("Upload Manifest", type=["txt"])

    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- RUN ----------------
if uploaded_image and uploaded_manifest:

    img_temp = tempfile.NamedTemporaryFile(delete=False)
    img_temp.write(uploaded_image.read())
    image_path = img_temp.name

    txt_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
    text = uploaded_manifest.read().decode("utf-8")

    with open(txt_temp.name, "w") as f:
        f.write(text)

    manifest_path = txt_temp.name

    if st.button("Run AI Analysis"):

        with st.spinner("Processing cargo inspection..."):

            result = pipeline.run(image_path, manifest_path)

            detections = result["detections"]
            declared_items = result["declared_items"]
            scores = result["scores"]
            decision = result["decision"]

            annotated = draw_boxes(image_path, detections)

        st.divider()

        left, center, right = st.columns([1.2,2,1.2])

        # ---------------- LEFT PANEL ----------------
        with left:

            st.markdown('<div class="card">', unsafe_allow_html=True)

            st.subheader("Declared Items")

            if declared_items:
                for item in declared_items:
                    st.write("•", item)
            else:
                st.write("No declared items")

            st.markdown("---")

            # ---------------- AI EXPLANATION ----------------

            st.subheader("AI Explanation")

            mismatch = round(scores["mismatch"],2)
            anomaly = round(scores["anomaly"],2)
            risk = round(scores["risk"],2)

            detected_names = [d.get("class", d.get("object")) for d in detections]

            if decision == "HIGH RISK":

                explanation = f"""
The AI system detected **{len(detected_names)} object(s)** in the X-ray scan: {", ".join(detected_names)}.

These detected objects **do not match the declared cargo manifest**, which produced a **high mismatch score ({mismatch})**.

The anomaly detection model also found **unusual cargo patterns** with an anomaly score of **{anomaly}**.

Therefore, this shipment is classified as **HIGH RISK** and should be **manually inspected by customs officers**.
"""

            elif decision == "SUSPICIOUS":

                explanation = f"""
The system detected **{len(detected_names)} object(s)** in the cargo scan: {", ".join(detected_names)}.

Some detected objects **partially differ from the declared cargo items**, resulting in a **moderate mismatch score ({mismatch})**.

The anomaly model also detected **minor irregularities** in the cargo structure.

Because of these inconsistencies, the shipment is marked as **SUSPICIOUS** and may require **additional inspection**.
"""

            else:

                explanation = f"""
The AI system detected **{len(detected_names)} object(s)** in the cargo scan: {", ".join(detected_names) if detected_names else "no objects"}.

The detected objects **match the declared manifest**, producing a **low mismatch score ({mismatch})**.

No unusual cargo patterns were detected by the anomaly model.

Therefore, the shipment is considered **SAFE** and no immediate risk was identified.
"""

            st.write(explanation)

            st.markdown('</div>', unsafe_allow_html=True)

        # ---------------- CENTER PANEL ----------------
        with center:

            st.markdown('<div class="card">', unsafe_allow_html=True)

            c1, c2 = st.columns(2)

            with c1:
                st.image(annotated, caption="Detected Objects")

            with c2:
                st.image(annotated, caption="GradCAM (placeholder)")

            st.markdown('</div>', unsafe_allow_html=True)

            # DETECTED ITEMS TABLE
            st.markdown('<div class="card">', unsafe_allow_html=True)

            st.subheader("Detected Items")

            if detections:
                df = pd.DataFrame([
                    {
                        "Object": d.get("class", d.get("object")),
                        "Confidence (%)": round(d.get("confidence",0)*100,2),
                        "BBox": d.get("bbox")
                    }
                    for d in detections
                ])
                st.dataframe(df, use_container_width=True)
            else:
                st.write("No detections")

            st.markdown('</div>', unsafe_allow_html=True)

            # METRICS
            st.subheader("AI Scores")

            c1,c2,c3,c4 = st.columns(4)

            c1.markdown(f'<div class="metric-card">Detection<br>{round(scores["detection"],2)}</div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="metric-card">Anomaly<br>{round(scores["anomaly"],2)}</div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="metric-card">Mismatch<br>{round(scores["mismatch"],2)}</div>', unsafe_allow_html=True)
            c4.markdown(f'<div class="metric-card">Risk<br>{round(scores["risk"],2)}</div>', unsafe_allow_html=True)

        # ---------------- RIGHT PANEL ----------------
        with right:

            st.markdown('<div class="card">', unsafe_allow_html=True)

            st.subheader("Risk Assessment")

            if decision == "SAFE":
                st.markdown('<div class="safe">SAFE</div>', unsafe_allow_html=True)

            elif decision == "SUSPICIOUS":
                st.markdown('<div class="warn">SUSPICIOUS</div>', unsafe_allow_html=True)

            else:
                st.markdown('<div class="danger">HIGH RISK</div>', unsafe_allow_html=True)

            st.markdown("---")

            st.subheader("Model Contributions")

            st.progress(float(scores["detection"]))
            st.write("YOLOv8")

            st.progress(float(scores["anomaly"]))
            st.write("ResNet50")

            st.progress(float(scores["mismatch"]))
            st.write("CLIP")

            st.markdown('</div>', unsafe_allow_html=True)