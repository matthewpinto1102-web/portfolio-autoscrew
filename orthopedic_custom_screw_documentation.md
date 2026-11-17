# Orthopedic Custom Screw Automation (`v6_screw.py`)

A Python automation script that interfaces with **SOLIDWORKS** via the COM API (`win32com`) to dynamically adjust parametric screw designs based on anatomical site, bone tissue quality, patient activity levels, and clinical health metrics.

## 🚀 Overview
Designing patient-specific orthopedic hardware requires tailoring geometric parameters to match bone density and biomechanical loads. This script automates that workflow by modifying model equations in real time, rebuilding the geometry, and saving a customized part file (`.SLDPRT`) into an organized output directory (`Screw_Parts`).

---

## 📋 Baseline Template & Requirements

### Required Active Document
The script does **not** hardcode or open a specific file path from your local system. Instead, it relies on whatever SOLIDWORKS part file you have **currently open and active** on your screen (`swApp.ActiveDoc`).

### Required Global Variables / Equations
For the script to execute successfully, your open SOLIDWORKS template must contain equation manager variables named exactly:
1. `"Radius"`
2. `"Pitch_Factor"`

---

## ⚙️ Features & Modifiers

### 1. Anatomical Site Base Radii (`SITE_BASE_RADIUS`)
Defines the starting radius (mm) for different skeletal locations:
* **Phalange:** $0.8 \text{ mm}$
* **Metacarpal:** $1.25 \text{ mm}$
* **Radius:** $1.5 \text{ mm}$
* **Humerus:** $2.0 \text{ mm}$
* **Tibia:** $2.25 \text{ mm}$
* **Femur:** $2.5 \text{ mm}$

### 2. Tissue Density Modifiers (`TISSUE_PF_MODIFIER`)
Adjusts pitch factor based on bone microarchitecture:
* **Cortical (Dense/Hard):** $+0.0$
* **Transition (Mixed):** $+0.1$
* **Cancellous (Spongy/Soft):** $+0.25$

### 3. Patient Health Delta Algorithm (`calculate_patient_pf_delta`)
Factors individual patient variables into structural changes:
* **Age Factor:** Adds incremental adjustments for patients over 30 years old.
* **Smoking Status:** Adds a penalty modifier ($+0.06$) for reduced osseointegration/bone quality.
* **Nutrition/Calcium Score:** Scales inversely based on a 1–10 health score.

---

## 🛠️ Prerequisites & Installation

* **OS:** Windows (Required for SOLIDWORKS COM API integration).
* **Software:** SOLIDWORKS installed and running with an active part document open.
* **Python Libraries:** `pywin32` (`win32com.client`).

Install the required library via terminal:
```bash
pip install pywin32
```

---

## 📂 Output Structure

The script automatically generates a local folder named **`Screw_Parts`** in the same directory as the script if it doesn't already exist. 

Generated files follow this naming convention:
$$\text{Patient\_Screw\_}\langle\text{Timestamp}\rangle\_\langle\text{Site}\rangle\_\langle\text{Tissue}\rangle\text{.SLDPRT}$$

*Example:* `Patient_Screw_1430_femur_cortical.SLDPRT`

---

## 💻 Example Usage

```python
# Generates a custom femur screw for a cortical tissue site, 
# athletic profile, age 44, smoker, with a nutrition score of 8.
v6_screw("femur", "cortical", "athlete", 44, True, 8)
```