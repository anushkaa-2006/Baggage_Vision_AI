#risk_mapping.py

def compute_risk(detection, anomaly, mismatch):
    w1, w2, w3 = 0.5, 0.3, 0.2
    risk = w1 * detection + w2 * anomaly + w3 * mismatch
    return round(risk, 2)


def get_decision(risk):
    if risk < 0.3:
        return "SAFE"
    elif risk < 0.6:
        return "SUSPICIOUS"
    else:
        return "HIGH RISK"