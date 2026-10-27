from mcp.server.fastmcp import FastMCP
import comtypes.client
import json

etabs = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
model = etabs.SapModel

server = FastMCP("etabs-mcp")

# ─────────────────────────────────────────────
# MODEL INFO & SETUP
# ─────────────────────────────────────────────

@server.tool()
def get_model_info() -> str:
    """Get basic model information: filename and current units"""
    return json.dumps({
        "filename": model.GetModelFilename(),
        "units": model.GetPresentUnits()
    })

@server.tool()
def create_new_model(units: int = 6) -> str:
    """
    Create a new blank ETABS model.
    units: 1=lb-in, 2=lb-ft, 3=kip-in, 4=kip-ft, 5=kN-mm, 6=kN-m, 7=kgf-m, 8=N-m
    """
    ret = model.InitializeNewModel(units)
    ret2 = model.File.NewBlank()
    return json.dumps({"status": "ok", "init": ret, "new_blank": ret2})

@server.tool()
def save_model(filepath: str) -> str:
    """Save the model to a file path (full path including .edb extension)"""
    ret = model.File.Save(filepath)
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def set_units(units: int) -> str:
    """
    Set present units.
    1=lb-in, 2=lb-ft, 3=kip-in, 4=kip-ft, 5=kN-mm, 6=kN-m, 7=kgf-m, 8=N-m
    """
    ret = model.SetPresentUnits(units)
    return json.dumps({"status": "ok", "ret": ret})

# ─────────────────────────────────────────────
# STORIES & GRIDS
# ─────────────────────────────────────────────

@server.tool()
def get_stories() -> str:
    """Get all story levels and elevations from ETABS"""
    result = model.Story.GetStories()
    # ETABS returns (num, names, elevations, ...) or just (names, elevations, ...)
    if isinstance(result[0], int):
        num, names, elevations = result[0], result[1], result[2]
    else:
        names, elevations = result[0], result[1]
        num = len(names)
    stories = [{"name": names[i], "elevation": elevations[i]} for i in range(num)]
    return json.dumps({"count": num, "stories": stories})

@server.tool()
def set_stories(story_names: list, story_heights: list) -> str:
    """
    Define stories by name and height (not elevation).
    story_names: list of story names e.g. ["Base","Story1","Story2"]
    story_heights: list of heights for each story e.g. [0, 3.0, 3.0]
    """
    n = len(story_names)
    ret = model.Story.SetStories_2(story_names, story_heights, [False]*n, [""]*n, [0.0]*n, [False]*n)
    return json.dumps({"status": "ok", "ret": ret})

# ─────────────────────────────────────────────
# MATERIALS
# ─────────────────────────────────────────────

@server.tool()
def get_materials() -> str:
    """Get all defined materials in the model"""
    result = model.PropMaterial.GetNameList()
    # Returns (num, names) or just (names,)
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    return json.dumps({"count": num, "materials": list(names)})

@server.tool()
def add_material_concrete(name: str, fc: float, is_lightweight: bool = False) -> str:
    """
    Add a concrete material.
    name: material name e.g. "C30"
    fc: compressive strength in current units e.g. 30000 kN/m2
    """
    ret = model.PropMaterial.SetMaterial(name, 2)
    ret2 = model.PropMaterial.SetOConcrete(name, fc, is_lightweight, 0, 2, 0.003, 0.0)
    return json.dumps({"status": "ok", "set_material": ret, "set_concrete": ret2})

@server.tool()
def add_material_steel(name: str, fy: float, fu: float) -> str:
    """
    Add a steel material.
    name: material name e.g. "S275"
    fy: yield strength in current units
    fu: ultimate strength in current units
    """
    ret = model.PropMaterial.SetMaterial(name, 1)
    ret2 = model.PropMaterial.SetOSteel(name, fy, fu, fy, fu, 1, 0.02, 0.1, 0.15, 0.0)
    return json.dumps({"status": "ok", "set_material": ret, "set_steel": ret2})

# ─────────────────────────────────────────────
# FRAME SECTIONS
# ─────────────────────────────────────────────

@server.tool()
def get_frame_sections() -> str:
    """Get all defined frame sections"""
    result = model.PropFrame.GetNameList()
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    return json.dumps({"count": num, "sections": list(names)})

@server.tool()
def add_rectangular_section(name: str, material: str, width: float, depth: float) -> str:
    """
    Add a rectangular frame section (e.g. concrete beam/column).
    name: section name, material: material name, width/depth in current units
    """
    ret = model.PropFrame.SetRectangle(name, material, depth, width)
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def add_steel_i_section(name: str, material: str, d: float, tw: float, bf: float, tf: float) -> str:
    """
    Add a steel I-section.
    d=total depth, tw=web thickness, bf=flange width, tf=flange thickness (current units)
    """
    ret = model.PropFrame.SetISection(name, material, d, tw, bf, tf, bf, tf)
    return json.dumps({"status": "ok", "ret": ret})

# ─────────────────────────────────────────────
# SHELL / SLAB SECTIONS
# ─────────────────────────────────────────────

@server.tool()
def get_shell_sections() -> str:
    """Get all defined shell/area sections"""
    result = model.PropArea.GetNameList()
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    return json.dumps({"count": num, "sections": list(names)})

@server.tool()
def add_slab_section(name: str, material: str, thickness: float) -> str:
    """
    Add a slab shell section.
    name: section name, material: material name, thickness in current units
    """
    ret = model.PropArea.SetSlab(name, 1, 1, material, thickness)
    return json.dumps({"status": "ok", "ret": ret})

# ─────────────────────────────────────────────
# JOINTS (NODES)
# ─────────────────────────────────────────────

@server.tool()
def add_joint(x: float, y: float, z: float) -> str:
    """Add a joint/node at given coordinates. Returns the joint name."""
    result = model.PointObj.AddCartesian(x, y, z, "")
    # Returns (ret, name) or (name, ret) depending on version
    if isinstance(result[0], int):
        ret, name = result[0], result[1]
    else:
        name, ret = result[0], result[1]
    return json.dumps({"status": "ok", "name": name, "ret": ret})

@server.tool()
def get_all_joints() -> str:
    """Get all joints/nodes in the model"""
    result = model.PointObj.GetNameList()
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    joints = []
    for n in names:
        coords = model.PointObj.GetCoordCartesian(n)
        x, y, z = coords[0], coords[1], coords[2]
        joints.append({"name": n, "x": x, "y": y, "z": z})
    return json.dumps({"count": num, "joints": joints})

@server.tool()
def assign_joint_restraint(joint_name: str, ux: bool, uy: bool, uz: bool, rx: bool, ry: bool, rz: bool) -> str:
    """
    Assign restraints to a joint (support conditions).
    True = restrained, False = free. Order: UX, UY, UZ, RX, RY, RZ
    """
    ret = model.PointObj.SetRestraint(joint_name, [ux, uy, uz, rx, ry, rz])
    return json.dumps({"status": "ok", "ret": ret})

# ─────────────────────────────────────────────
# FRAME OBJECTS (BEAMS & COLUMNS)
# ─────────────────────────────────────────────

@server.tool()
def add_frame(xi: float, yi: float, zi: float, xj: float, yj: float, zj: float, section: str) -> str:
    """
    Add a frame element (beam or column) between two points.
    Provide start (i) and end (j) coordinates and section name.
    """
    result = model.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj, "", section)
    if isinstance(result[0], int):
        ret, name = result[0], result[1]
    else:
        name, ret = result[0], result[1]
    return json.dumps({"status": "ok", "name": name, "ret": ret})

@server.tool()
def get_all_frames() -> str:
    """Get all frame objects in the model"""
    result = model.FrameObj.GetNameList()
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    frames = []
    for n in names:
        sec_result = model.FrameObj.GetSection(n)
        sec = sec_result[0]
        frames.append({"name": n, "section": sec})
    return json.dumps({"count": num, "frames": frames})

# ─────────────────────────────────────────────
# AREA OBJECTS (SLABS / WALLS)
# ─────────────────────────────────────────────

@server.tool()
def add_area_by_corners(x_coords: list, y_coords: list, z_coords: list, section: str) -> str:
    """
    Add an area object (slab/wall) by corner coordinates.
    x_coords, y_coords, z_coords: lists of corner coordinates (3 or 4 points)
    section: shell section name
    """
    n = len(x_coords)
    result = model.AreaObj.AddByCoord(n, x_coords, y_coords, z_coords, "", section)
    if isinstance(result[0], int):
        ret, name = result[0], result[1]
    else:
        name, ret = result[0], result[1]
    return json.dumps({"status": "ok", "name": name, "ret": ret})

@server.tool()
def get_all_areas() -> str:
    """Get all area objects in the model"""
    result = model.AreaObj.GetNameList()
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    return json.dumps({"count": num, "areas": list(names)})

# ─────────────────────────────────────────────
# LOAD PATTERNS & CASES
# ─────────────────────────────────────────────

@server.tool()
def get_load_patterns() -> str:
    """Get all defined load patterns"""
    result = model.LoadPatterns.GetNameList()
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    return json.dumps({"count": num, "patterns": list(names)})

@server.tool()
def add_load_pattern(name: str, load_type: int = 1, self_weight_multiplier: float = 0.0) -> str:
    """
    Add a load pattern.
    load_type: 1=Dead, 2=SuperDead, 3=Live, 4=ReducedLive, 5=Wind, 6=Seismic, 8=Snow
    self_weight_multiplier: 1.0 to include self weight
    """
    ret = model.LoadPatterns.Add(name, load_type, self_weight_multiplier, True)
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def get_load_cases() -> str:
    """Get all defined load cases"""
    result = model.LoadCases.GetNameList()
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    return json.dumps({"count": num, "cases": list(names)})

@server.tool()
def add_static_linear_load_case(name: str, load_pattern_names: list, scale_factors: list) -> str:
    """
    Add a static linear load case.
    load_pattern_names: list of pattern names e.g. ["Dead","Live"]
    scale_factors: list of scale factors e.g. [1.0, 1.5]
    """
    ret = model.LoadCases.StaticLinear.SetCase(name)
    n = len(load_pattern_names)
    ret2 = model.LoadCases.StaticLinear.SetLoads(name, n, ["Load"]*n, load_pattern_names, scale_factors)
    return json.dumps({"status": "ok", "set_case": ret, "set_loads": ret2})

# ─────────────────────────────────────────────
# FRAME LOADS
# ─────────────────────────────────────────────

@server.tool()
def assign_frame_distributed_load(frame_name: str, load_pattern: str, load_value: float, direction: int = 6) -> str:
    """
    Assign a uniform distributed load to a frame.
    direction: 1=Local1, 2=Local2, 3=Local3, 4=X, 5=Y, 6=Z (gravity)
    load_value: load magnitude in current units (use negative for downward in Z)
    """
    ret = model.FrameObj.SetLoadDistributed(frame_name, load_pattern, 1, direction, 0.0, 1.0, load_value, load_value)
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def assign_joint_force(joint_name: str, load_pattern: str, fx: float = 0, fy: float = 0, fz: float = 0, mx: float = 0, my: float = 0, mz: float = 0) -> str:
    """Assign point force/moment to a joint"""
    ret = model.PointObj.SetLoadForce(joint_name, load_pattern, [fx, fy, fz, mx, my, mz], False)
    return json.dumps({"status": "ok", "ret": ret})

# ─────────────────────────────────────────────
# ANALYSIS
# ─────────────────────────────────────────────

@server.tool()
def run_analysis() -> str:
    """Run the ETABS analysis"""
    ret = model.Analyze.RunAnalysis()
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def get_available_results_cases() -> str:
    """Get list of load cases available for results extraction"""
    result = model.Results.Setup.GetCaseSelectedForOutput()
    if isinstance(result[0], int):
        num, cases = result[0], result[1]
    else:
        cases = result[0]
        num = len(cases)
    return json.dumps({"count": num, "cases": list(cases)})

@server.tool()
def select_results_case(case_name: str) -> str:
    """Select a load case for results output"""
    model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = model.Results.Setup.SetCaseSelectedForOutput(case_name)
    return json.dumps({"status": "ok", "ret": ret})

# ─────────────────────────────────────────────
# RESULTS
# ─────────────────────────────────────────────

@server.tool()
def get_joint_displacements(joint_name: str, case_name: str) -> str:
    """
    Get joint displacements for a specific joint and load case.
    Returns UX, UY, UZ, RX, RY, RZ.
    """
    model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    model.Results.Setup.SetCaseSelectedForOutput(case_name)
    result = model.Results.JointDispl(joint_name, 0)
    num = result[0]
    load, ux, uy, uz, rx, ry, rz = result[3], result[6], result[7], result[8], result[9], result[10], result[11]
    results = []
    for i in range(num):
        results.append({"load_case": load[i], "UX": ux[i], "UY": uy[i], "UZ": uz[i], "RX": rx[i], "RY": ry[i], "RZ": rz[i]})
    return json.dumps({"joint": joint_name, "results": results})

@server.tool()
def get_frame_forces(frame_name: str, case_name: str) -> str:
    """
    Get frame internal forces for a specific frame and load case.
    Returns P, V2, V3, T, M2, M3 along the frame.
    """
    model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    model.Results.Setup.SetCaseSelectedForOutput(case_name)
    result = model.Results.FrameForce(frame_name, 0)
    num = result[0]
    load, loc, p, v2, v3, t, m2, m3 = result[3], result[4], result[5], result[6], result[7], result[8], result[9], result[10]
    results = []
    for i in range(num):
        results.append({"load_case": load[i], "loc": loc[i], "P": p[i], "V2": v2[i], "V3": v3[i], "T": t[i], "M2": m2[i], "M3": m3[i]})
    return json.dumps({"frame": frame_name, "results": results})

@server.tool()
def get_joint_reactions(joint_name: str, case_name: str) -> str:
    """Get support reactions at a joint for a given load case"""
    model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    model.Results.Setup.SetCaseSelectedForOutput(case_name)
    result = model.Results.JointReact(joint_name, 0)
    num = result[0]
    load, fx, fy, fz, mx, my, mz = result[3], result[6], result[7], result[8], result[9], result[10], result[11]
    results = []
    for i in range(num):
        results.append({"load_case": load[i], "FX": fx[i], "FY": fy[i], "FZ": fz[i], "MX": mx[i], "MY": my[i], "MZ": mz[i]})
    return json.dumps({"joint": joint_name, "results": results})

@server.tool()
def get_story_drifts(case_name: str) -> str:
    """Get story drifts for a given load case"""
    model.Results.Setup.DeselectAllCasesAndCombosForOutput()
    model.Results.Setup.SetCaseSelectedForOutput(case_name)
    result = model.Results.StoryDrifts()
    num = result[0]
    story, load, direction, drift = result[1], result[2], result[5], result[6]
    results = []
    for i in range(num):
        results.append({"story": story[i], "load_case": load[i], "direction": direction[i], "drift": drift[i]})
    return json.dumps({"results": results})

# ─────────────────────────────────────────────
# COMBINATIONS
# ─────────────────────────────────────────────

@server.tool()
def add_load_combination(name: str, combo_type: int, case_names: list, scale_factors: list) -> str:
    """
    Add a load combination.
    combo_type: 0=Linear Additive, 1=Envelope, 2=Absolute, 3=SRSS
    case_names: list of load case or combo names
    scale_factors: corresponding scale factors
    """
    ret = model.RespCombo.Add(name, combo_type)
    for cn, sf in zip(case_names, scale_factors):
        model.RespCombo.SetCaseList(name, 0, cn, sf)
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def get_load_combinations() -> str:
    """Get all defined load combinations"""
    result = model.RespCombo.GetNameList()
    if isinstance(result[0], int):
        num, names = result[0], result[1]
    else:
        names = result[0]
        num = len(names)
    return json.dumps({"count": num, "combinations": list(names)})

# ─────────────────────────────────────────────
# DESIGN
# ─────────────────────────────────────────────

@server.tool()
def run_steel_design() -> str:
    """Run steel frame design"""
    ret = model.DesignSteel.StartDesign()
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def run_concrete_design() -> str:
    """Run concrete frame design"""
    ret = model.DesignConcrete.StartDesign()
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def get_steel_design_results(frame_name: str) -> str:
    """Get steel design D/C ratio for a frame element"""
    result = model.DesignSteel.GetSummaryResults(frame_name)
    dc_ratio, combo, status = result[0], result[1], result[2]
    return json.dumps({"frame": frame_name, "dc_ratio": dc_ratio, "governing_combo": combo, "status": status})

# ─────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────

@server.tool()
def refresh_view() -> str:
    """Refresh the ETABS view"""
    ret = model.View.RefreshView(0, False)
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def lock_model() -> str:
    """Lock the model (required before running analysis)"""
    ret = model.SetModelIsLocked(True)
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def unlock_model() -> str:
    """Unlock the model (required to make changes after analysis)"""
    ret = model.SetModelIsLocked(False)
    return json.dumps({"status": "ok", "ret": ret})

@server.tool()
def get_model_is_locked() -> str:
    """Check if the model is currently locked"""
    result = model.GetModelIsLocked()
    locked = result[0] if isinstance(result, (list, tuple)) else result
    return json.dumps({"locked": locked})

# ─────────────────────────────────────────────

if __name__ == "__main__":
    server.run(transport="stdio")