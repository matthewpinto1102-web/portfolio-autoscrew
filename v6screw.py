import win32com.client
import os
from datetime import datetime

script_dir = os.path.dirname(os.path.abspath(__file__))
output_folder_name = "Screw_Parts"
output_dir = os.path.join(script_dir, output_folder_name)
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
    print(f"Created new folder: {output_folder_name}")
SITE_BASE_RADIUS = {
    "phalange": 0.8,
    "metacarpal": 1.25,
    "radius": 1.5,
    "humerus": 2.0,
    "femur": 2.5,
    "tibia": 2.25
}
TISSUE_PF_MODIFIER = {
    "cortical": 0.0,      # Dense/Hard
    "transition": 0.1,    # Mixed
    "cancellous": 0.25    # Spongy/Soft
}
ACTIVITY_PF_MODIFIER = {
    "sedentary": 0.05,
    "moderate": 0.0,
    "athlete": -0.05      # Tightens thread to leave more metal in the core
}
def calculate_patient_pf_delta(age, is_smoker, health_score):
    """
    age: int
    is_smoker: boolean
    health_score: 1 to 10 (10 being perfect calcium/nutrition)
    """
    delta = 0
    if age > 30:
        delta += (age - 30) * 0.001
    if is_smoker:
        delta += 0.06
    delta -= (health_score - 5) * 0.01
    return delta

def v6_screw(site, tissue, activity, age, smoking, nutrition):
    try:
        swApp = win32com.client.dynamic.Dispatch("SldWorks.Application")
        model = swApp.ActiveDoc
        eqMgr = model.GetEquationMgr
        found1 = False
        found2 = False
        base_pf = 0.25 # The "Ideal" Cortical PF
        tissue_mod = TISSUE_PF_MODIFIER.get(tissue, 0.1)
        activity_mod = ACTIVITY_PF_MODIFIER.get(activity, 0.0)
        new_radius=SITE_BASE_RADIUS.get(site,1.6)
        patient_delta = calculate_patient_pf_delta(age, smoking, nutrition)
        new_pitch_factor = base_pf + tissue_mod + activity_mod + patient_delta
        #Adjustments
        if new_radius>4.0:
            new_radius=4.0
        if new_radius<.5:
            new_radius=.5  
        if new_pitch_factor>.75:
            new_pitch_factor=.75
        if new_pitch_factor<.15:
            new_pitch_factor=.15
        for i in range(eqMgr.GetCount):
            if '"Radius"' in eqMgr.Equation(i):
                print(f"Found Master Radius at Index {i}. Updating to {new_radius}mm...")
                eqMgr.Equation(i, f'"Radius" = {new_radius}mm')
                found1 = True
                break
        for j in range(eqMgr.GetCount):
            if '"Pitch_Factor"' in eqMgr.Equation(j):
                print(f"Found PF at index {j}. Updating to {new_pitch_factor}")
                eqMgr.Equation(j, f'"Pitch_Factor"={new_pitch_factor}')
                found2=True
                break
        if found1 and found2:
            eqMgr.EvaluateAll
            model.ForceRebuild3(True)
            print("Successful update! Check SOLIDWORKS model.")
            script_dir = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.join(script_dir, "Screw_Parts")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            timestamp = datetime.now().strftime("%H%M")
            new_filename = f'Patient_Screw_{timestamp}_{site}_{tissue}.SLDPRT'
            full_path = os.path.join(output_dir, new_filename)
            full_path = full_path.replace('/', '\\')
            save_options = 1 + 2 
            success = model.SaveAs3(full_path, 0, save_options)
            if success == 0:
                print(f"Saved to: {full_path}")
            else:
                print(f"Save Failed, errorr code: {success}")
    except Exception as e:
        print(f"Error: {e}")
        
        
        
        
        
        
        
v6_screw("femur","cortical","athlete",44,True,8)