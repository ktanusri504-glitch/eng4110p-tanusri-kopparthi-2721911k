#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from part import *
from material import *
from section import *
from assembly import *
from step import *
from interaction import *
from load import *
from mesh import *
from optimization import *
from job import *
from sketch import *
from visualization import *
from connectorBehavior import *
import numpy as np
import math
import zipfile
import odbAccess

import os


OUTPUT_DIR = r"H:/WIND TURBINE"
if not os.path.isdir(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

xlsx_name = os.path.join(OUTPUT_DIR, 'mc_wind_results_by_height.xlsx')
print("Excel will be written to:", os.path.abspath(xlsx_name))



from abaqusConstants import *
from regionToolset import Region

try:
    integer_types = (int, long, np.integer)
except NameError:
    integer_types = (int, np.integer)


# ------------------------------------------------------------
# helpers: lightweight .xlsx writer
# ------------------------------------------------------------
def _xml_escape(text):
    s = str(text)
    s = s.replace('&', '&amp;')
    s = s.replace('<', '&lt;')
    s = s.replace('>', '&gt;')
    s = s.replace('"', '&quot;')
    s = s.replace("'", '&apos;')
    return s

def _excel_col(idx):
    name = ''
    n = idx
    while n > 0:
        n, rem = divmod(n - 1, 26)
        name = chr(65 + rem) + name
    return name

def _is_number(v):
    return isinstance(v, integer_types) or isinstance(v, (float, np.floating))

def _fmt_num(v):
    if isinstance(v, integer_types):
        return str(int(v))
    return ('%.12g' % float(v))

def write_xlsx(path, sheet_name, headers, data_rows):
    all_rows = [headers] + data_rows

    row_xml = []
    for r_idx in range(len(all_rows)):
        row = all_rows[r_idx]
        excel_row = r_idx + 1
        cells = []
        for c_idx in range(len(row)):
            v = row[c_idx]
            ref = '%s%d' % (_excel_col(c_idx + 1), excel_row)
            if v is None:
                cells.append('<c r="%s"/>' % ref)
            elif _is_number(v):
                cells.append('<c r="%s"><v>%s</v></c>' % (ref, _fmt_num(v)))
            else:
                cells.append(
                    '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, _xml_escape(v))
                )
        row_xml.append('<row r="%d">%s</row>' % (excel_row, ''.join(cells)))

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>%s</sheetData>'
        '</worksheet>'
    ) % ''.join(row_xml)

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )

    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '</Relationships>'
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    ) % _xml_escape(sheet_name)

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        '</Relationships>'
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )

    z = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
    z.writestr('[Content_Types].xml', content_types_xml)
    z.writestr('_rels/.rels', rels_xml)
    z.writestr('xl/workbook.xml', workbook_xml)
    z.writestr('xl/_rels/workbook.xml.rels', workbook_rels_xml)
    z.writestr('xl/worksheets/sheet1.xml', sheet_xml)
    z.writestr('xl/styles.xml', styles_xml)
    z.close()

def build_avg_by_node(field_values, odb_instance_name, scalar_fn):
    acc = {}
    cnt = {}
    iname = odb_instance_name.upper()
    for v in field_values:
        if (not hasattr(v, 'instance')) or (v.instance is None):
            continue
        if (not hasattr(v, 'nodeLabel')) or (v.nodeLabel is None):
            continue
        if v.instance.name.upper() != iname:
            continue

        node_label = v.nodeLabel
        sval = scalar_fn(v)
        if sval is None:
            continue
        if node_label in acc:
            acc[node_label] += sval
            cnt[node_label] += 1
        else:
            acc[node_label] = sval
            cnt[node_label] = 1

    out = {}
    for k in acc.keys():
        out[k] = acc[k] / float(cnt[k])
    return out

# ------------------------------------------------------------
# 1) Geometry: tapered shell tower 
# ------------------------------------------------------------
H_mm = 80000.0
R_bot_mm = 2100.0
R_top_mm = 1200.0

mdb.models['Model-1'].ConstrainedSketch(name='__profile__', sheetSize=200000.0)
s = mdb.models['Model-1'].sketches['__profile__']

# Axis of revolution
s.ConstructionLine(point1=(0.0, -100000.0), point2=(0.0, 100000.0))

# Generator line of the conical shell
s.Line(point1=(R_bot_mm, 0.0), point2=(R_top_mm, H_mm))

mdb.models['Model-1'].Part(
    dimensionality=THREE_D,
    name='Tower',
    type=DEFORMABLE_BODY
)
mdb.models['Model-1'].parts['Tower'].BaseShellRevolve(
    angle=360.0,
    flipRevolveDirection=OFF,
    sketch=s
)

del mdb.models['Model-1'].sketches['__profile__']

p = mdb.models['Model-1'].parts['Tower']

# ------------------------------------------------------------
# 2) Material + section (Steel, 30 mm shell)
# ------------------------------------------------------------
mdb.models['Model-1'].Material(name='Steel')
mdb.models['Model-1'].materials['Steel'].Density(table=((7.85e-09,),))     # tonne/mm^3
mdb.models['Model-1'].materials['Steel'].Elastic(table=((200000.0, 0.3),)) # MPa

mdb.models['Model-1'].HomogeneousShellSection(
    name='Tower Section',
    material='Steel',
    thickness=30.0,
    idealization=NO_IDEALIZATION,
    integrationRule=SIMPSON,
    numIntPts=5,
    thicknessType=UNIFORM
)

# Masks become fragile; assign to all faces
p.SectionAssignment(
    region=Region(faces=p.faces),
    sectionName='Tower Section',
    offset=0.0,
    offsetType=BOTTOM_SURFACE,
    thicknessAssignment=FROM_SECTION
)

# ------------------------------------------------------------
# 3) Assembly 
# ------------------------------------------------------------
mdb.models['Model-1'].rootAssembly.DatumCsysByDefault(CARTESIAN)
mdb.models['Model-1'].rootAssembly.Instance(
    dependent=ON, name='Tower-1', part=p
)

m = mdb.models['Model-1']
a = m.rootAssembly
inst = a.instances['Tower-1']

BIG = 1.0e12

# ------------------------------------------------------------
# 4) Step: Static ONLY
# ------------------------------------------------------------
STEP_NAME = 'Static-step'
m.StaticStep(name=STEP_NAME, previous='Initial', nlgeom=ON)

# ------------------------------------------------------------
# 5) Top reference point + rigid body constraint
# ------------------------------------------------------------
a.ReferencePoint(point=(0.0, H_mm, 0.0))
rp_key = sorted(a.referencePoints.keys())[-1]
rp = a.referencePoints[rp_key]
a.Set(name='TopRP', referencePoints=(rp,))

top_edges = inst.edges.getByBoundingBox(
    xMin=-BIG, xMax=BIG,
    yMin=H_mm - 1.0, yMax=H_mm + 1.0,
    zMin=-BIG, zMax=BIG
)

print('Top edge count = %d' % len(top_edges))

m.RigidBody(
    name='Tower-top-RBC',
    refPointRegion=Region(referencePoints=(rp,)),
    tieRegion=Region(edges=top_edges)
)

a.engineeringFeatures.PointMassInertia(
    name='Nacelle mass',
    region=Region(referencePoints=(rp,)),
    mass=82.0,
    alpha=0.0,
    composite=0.0
)

a.engineeringFeatures.PointMassInertia(
    name='Rotor mass',
    region=Region(referencePoints=(rp,)),
    mass=60.0,
    alpha=0.0,
    composite=0.0
)


# ------------------------------------------------------------
# 6) Boundary condition: fixed base (Static ONLY)
# ------------------------------------------------------------
base_edges = inst.edges.getByBoundingBox(
    xMin=-BIG, xMax=BIG,
    yMin=-1.0, yMax=1.0,
    zMin=-BIG, zMax=BIG
)

print('Base edge count = %d' % len(base_edges))

m.DisplacementBC(
    name='FixedBase-Static',
    createStepName='Static-step',
    region=Region(edges=base_edges),
    u1=0.0, u2=0.0, u3=0.0, ur1=0.0, ur2=0.0, ur3=0.0
)

# ------------------------------------------------------------
# 7) Loads in Static step (Gravity + HEIGHT-VARYING wind via FIELD)
# ------------------------------------------------------------
m.Gravity(name='Gravity', createStepName='Static-step', comp2=-9810.0)  # mm/s^2

# Surface for wind
a.Surface(name='TowerSurf', side1Faces=inst.faces)

# ---- Wind/load parameters used by the field definition
rho_air = 1.225
Cd = 1.2
Pa_to_Nmm2 = 1e-6             # 1 Pa = 1e-6 N/mm^2

yref_mm = 10000.0             # 10 m reference height
alpha = 0.15
expn = 2.0 * alpha

USE_PI_CORRECTION = True      # recommended for full-cylinder traction with fixed +X direction

def p_ref_from_speed_Nmm2(Vref_ms):
    p_pa = 0.5 * rho_air * Cd * (Vref_ms ** 2)  # Pa
    if USE_PI_CORRECTION:
        p_pa = p_pa / np.pi
    return p_pa * Pa_to_Nmm2                     # N/mm^2

# Initial placeholder speed only for creating the field.
# Actual case-by-case wind speeds are assigned in Section 10.
Vref_init = 10.0
p0 = p_ref_from_speed_Nmm2(Vref_init)  # N/mm^2 at 10 m

print('---------------------------------------------')
print('Height-varying Wind Pressure (q_wind)')
print('Reference height (yref) = %.1f mm' % yref_mm)
print('Shear exponent (alpha)  = %.3f' % alpha)
print('p0_init @ 10m           = %.3e N/mm^2' % p0)
print('---------------------------------------------')

# Print initial placeholder values every 10 m
for h in range(0, int(H_mm / 1000) + 10, 10):
    h_mm = h * 1000.0
    q_val = p0 * (((h_mm + 1.0) / yref_mm) ** expn)
    print('Height = %2d m  ->  q_init = %.3e N/mm^2  (%.1f Pa)'
          % (h, q_val, q_val * 1e6))

print('---------------------------------------------')

# Wind direction (+X)
wind_dir = ((0.0, 0.0, 0.0), (1.0, 0.0, 0.0))

expr = ('(%0.12e) * (((Y+1.0)/%0.12e)**(%0.12e))' % (p0, yref_mm, expn))

m.ExpressionField(
    name='WindPressField',
    localCsys=None,
    description='p(Y) = p0*((Y/yref)^(2a)) [N/mm^2], with Y in mm',
    expression=expr
)

m.SurfaceTraction(
    name='Wind',
    createStepName='Static-step',
    region=a.surfaces['TowerSurf'],
    magnitude=1.0,
    directionVector=wind_dir,
    traction=GENERAL,
    distributionType=FIELD,
    field='WindPressField'
)

print('Initial wind field applied:')
print('  Vref_init @10m = %.3f m/s' % Vref_init)
print('  p0_init @10m   = %.3e N/mm^2' % p0)
print('  expression =', expr)


# ------------------------------------------------------------
# 8) Meshing
# ------------------------------------------------------------
p.seedPart(size=500.0, deviationFactor=0.1, minSizeFactor=0.1)

p.setElementType(
    elemTypes=(
        ElemType(elemCode=S8R, elemLibrary=STANDARD),
        ElemType(elemCode=STRI65, elemLibrary=STANDARD)
    ),
    regions=(p.faces,)
)

p.generateMesh()
a.regenerate()
inst = a.instances['Tower-1']

# ------------------------------------------------------------
# 9) Precompute node rings every 10 m for reporting
# ------------------------------------------------------------
height_levels_mm = [float(h * 1000) for h in range(0, int(H_mm / 1000) + 1, 10)]  # 0..80m
ring_tol_mm = 250.0

ring_labels_by_height = {}
for y_target in height_levels_mm:
    labels = []
    for nd in inst.nodes:
        if abs(nd.coordinates[1] - y_target) <= ring_tol_mm:
            labels.append(nd.label)
    ring_labels_by_height[y_target] = labels
    print('Height %.1f m -> %d nodes selected for averaging' % (y_target / 1000.0, len(labels)))

# ------------------------------------------------------------
# 10) Two-run study: NO_GUST vs ALL_GUST with range included
#     and most samples clustered near the mean
# ------------------------------------------------------------
N_MC = 100
SEED = 20260318

# ------------------------------------------------------------
# Site statistics from dataset (01 Jan 2024 to 18 Mar 2026)
# Data values were in mph, so convert to m/s here
# ------------------------------------------------------------
mph_to_ms = 0.44704

WS_MIN = 3.0 * mph_to_ms
WS_MAX = 51.4 * mph_to_ms
WS_MEDIAN = 14.1 * mph_to_ms
WS_MEAN = 15.223267 * mph_to_ms
WS_SD = 6.3752856 * mph_to_ms

GUST_MIN = 6.2 * mph_to_ms
GUST_MAX = 80.9 * mph_to_ms
GUST_MEDIAN = 23.2 * mph_to_ms
GUST_MEAN = 24.090718 * mph_to_ms
GUST_SD = 10.1203 * mph_to_ms

np.random.seed(SEED)

def truncated_normal_samples(n_needed, mu, sigma, vmin, vmax):
    out = []
    while len(out) < n_needed:
        v_try = np.random.normal(mu, sigma)
        if vmin <= v_try <= vmax:
            out.append(v_try)
    return out

def build_samples_with_range(n_total, vmin, vmedian, vmax, mu, sigma):
    fixed = [vmin, vmedian, vmax]
    remaining = n_total - len(fixed)
    sampled = truncated_normal_samples(remaining, mu, sigma, vmin, vmax)
    all_vals = fixed + sampled
    np.random.shuffle(all_vals)
    return np.array(all_vals)

# Run A: no gusts -> use wind speed distribution
v_mean_samples = build_samples_with_range(
    N_MC, WS_MIN, WS_MEDIAN, WS_MAX, WS_MEAN, WS_SD
)
gust_flags_no = np.zeros(N_MC, dtype=int)
v_peak_no = v_mean_samples.copy()

# Run B: all gusts -> use gust distribution directly as peak speed
v_peak_all = build_samples_with_range(
    N_MC, GUST_MIN, GUST_MEDIAN, GUST_MAX, GUST_MEAN, GUST_SD
)
gust_flags_all = np.ones(N_MC, dtype=int)

def append_case_rows(results_rows, case_id, v_mean, v_peak, gust_flag, p0_case,
                     u_by_node, s_by_node, e_by_node, run_status):

    for y_target in height_levels_mm:
        height_m = y_target / 1000.0
        v_mean_h = v_mean * (((y_target + 1.0) / yref_mm) ** alpha)
        v_peak_h = v_peak * (((y_target + 1.0) / yref_mm) ** alpha)

        p_h = p0_case * (((y_target + 1.0) / yref_mm) ** expn)
        p_h_pa = p_h * 1e6

        labels = ring_labels_by_height[y_target]

        if run_status == 'COMPLETED':
            u_vals = [u_by_node[lbl] for lbl in labels if lbl in u_by_node]
            s_vals = [s_by_node[lbl] for lbl in labels if lbl in s_by_node]
            e_vals = [e_by_node[lbl] for lbl in labels if lbl in e_by_node]

            u_avg = (sum(u_vals) / float(len(u_vals))) if len(u_vals) > 0 else None
            s_avg = (sum(s_vals) / float(len(s_vals))) if len(s_vals) > 0 else None
            e_avg = (sum(e_vals) / float(len(e_vals))) if len(e_vals) > 0 else None
        else:
            u_avg = None
            s_avg = None
            e_avg = None

        results_rows.append([
            case_id,
            height_m,
            v_mean,
            v_mean_h,
            gust_flag,
            v_peak,
            v_peak_h,
            p_h_pa,
            u_avg,
            s_avg,
            e_avg,
            run_status
        ])

def run_batch(run_tag, v_mean_arr, v_peak_arr, gust_flags, out_xlsx, case_summary_rows):
    results_rows = []

    for idx in range(N_MC):
        case_id = idx + 1
        v_mean = float(v_mean_arr[idx])
        v_peak = float(v_peak_arr[idx])
        gust_flag = int(gust_flags[idx])

        p0_case = p_ref_from_speed_Nmm2(v_peak)
        expr_case = '(%0.12e) * (((Y+1.0)/%0.12e)**(%0.12e))' % (p0_case, yref_mm, expn)
        m.analyticalFields['WindPressField'].setValues(expression=expr_case)

        if run_tag == 'NO_GUST':
            job_name = 'WIND_ONLY_%03d' % case_id
        else:
            job_name = 'WITH_GUSTS_%03d' % case_id

        if job_name in mdb.jobs.keys():
            del mdb.jobs[job_name]

        myJob = mdb.Job(
            name=job_name,
            model='Model-1',
            description='%s case %d: v_mean=%.3f m/s, v_peak=%.3f m/s, gust=%d'
                        % (run_tag, case_id, v_mean, v_peak, gust_flag),
            type=ANALYSIS
        )

        print('Running %s | v_mean=%.3f | v_peak=%.3f | gust=%d'
              % (job_name, v_mean, v_peak, gust_flag))

        myJob.submit(consistencyChecking=OFF)
        myJob.waitForCompletion()

        status = mdb.jobs[job_name].status
        if status != COMPLETED:
            print('WARNING: %s did not complete (status=%s). Writing empty rows.' % (job_name, status))
            append_case_rows(results_rows, case_id, v_mean, v_peak, gust_flag, p0_case,
                             {}, {}, {}, 'FAILED')

            case_summary_rows.append([
                run_tag,
                case_id,
                v_mean,
                v_peak,
                gust_flag,
                'FAILED',
                None,
                None,
                None
            ])
            continue

        odb_path = job_name + '.odb'
        if not os.path.exists(odb_path):
            print('WARNING: ODB missing for %s. Writing empty rows.' % job_name)
            append_case_rows(results_rows, case_id, v_mean, v_peak, gust_flag, p0_case,
                             {}, {}, {}, 'FAILED')

            case_summary_rows.append([
                run_tag,
                case_id,
                v_mean,
                v_peak,
                gust_flag,
                'FAILED',
                None,
                None,
                None
            ])
            continue

        odb = odbAccess.openOdb(path=odb_path, readOnly=True)

        step_names = list(odb.steps.keys())
        if len(step_names) == 0:
            print('WARNING: No steps in ODB for %s. Writing empty rows.' % job_name)
            odb.close()
            append_case_rows(results_rows, case_id, v_mean, v_peak, gust_flag, p0_case,
                             {}, {}, {}, 'FAILED')

            case_summary_rows.append([
                run_tag,
                case_id,
                v_mean,
                v_peak,
                gust_flag,
                'FAILED',
                None,
                None,
                None
            ])
            continue

        if STEP_NAME in step_names:
            step_odb = odb.steps[STEP_NAME]
        else:
            fallback_step = step_names[-1]
            print('WARNING: Step %s not found in %s. Using %s.'
                  % (STEP_NAME, job_name, fallback_step))
            step_odb = odb.steps[fallback_step]

        if len(step_odb.frames) == 0:
            print('WARNING: No frames in step for %s. Writing empty rows.' % job_name)
            odb.close()
            append_case_rows(results_rows, case_id, v_mean, v_peak, gust_flag, p0_case,
                             {}, {}, {}, 'FAILED')

            case_summary_rows.append([
                run_tag,
                case_id,
                v_mean,
                v_peak,
                gust_flag,
                'FAILED',
                None,
                None,
                None
            ])
            continue

        last_frame = step_odb.frames[-1]

        odb_inst_name = None
        for nm in odb.rootAssembly.instances.keys():
            if nm.upper() == inst.name.upper():
                odb_inst_name = nm
                break
        if odb_inst_name is None:
            odb_inst_name = list(odb.rootAssembly.instances.keys())[0]

        u_by_node = {}
        if 'U' in last_frame.fieldOutputs.keys():
            for v in last_frame.fieldOutputs['U'].values:
                if (not hasattr(v, 'instance')) or (v.instance is None):
                    continue
                if (not hasattr(v, 'nodeLabel')) or (v.nodeLabel is None):
                    continue
                if v.instance.name.upper() != odb_inst_name.upper():
                    continue
                ux, uy, uz = v.data
                u_by_node[v.nodeLabel] = math.sqrt(ux * ux + uy * uy + uz * uz)

        s_by_node = {}
        if 'S' in last_frame.fieldOutputs.keys():
            s_field = last_frame.fieldOutputs['S'].getSubset(position=ELEMENT_NODAL)
            s_by_node = build_avg_by_node(
                s_field.values,
                odb_inst_name,
                lambda vv: vv.mises
            )

        strain_key = None
        if 'LE' in last_frame.fieldOutputs.keys():
            strain_key = 'LE'
        elif 'E' in last_frame.fieldOutputs.keys():
            strain_key = 'E'

        e_by_node = {}
        if strain_key is not None:
            e_field = last_frame.fieldOutputs[strain_key].getSubset(position=ELEMENT_NODAL)

            def strain_scalar(vv):
                d = vv.data
                try:
                    total = 0.0
                    for c in d:
                        total += c * c
                    return math.sqrt(total)
                except:
                    try:
                        return abs(float(d))
                    except:
                        return None

            e_by_node = build_avg_by_node(
                e_field.values,
                odb_inst_name,
                strain_scalar
            )

        max_u_case = max(u_by_node.values()) if len(u_by_node) > 0 else None
        max_s_case = max(s_by_node.values()) if len(s_by_node) > 0 else None
        max_e_case = max(e_by_node.values()) if len(e_by_node) > 0 else None

        append_case_rows(results_rows, case_id, v_mean, v_peak, gust_flag, p0_case,
                         u_by_node, s_by_node, e_by_node, 'COMPLETED')

        case_summary_rows.append([
            run_tag,
            case_id,
            v_mean,
            v_peak,
            gust_flag,
            'COMPLETED',
            max_u_case,
            max_s_case,
            max_e_case
        ])

        odb.close()

    headers = [
        'Case_ID',
        'Height_m',
        'Wind_Mean_10m_m/s',
        'Wind_Mean_at_Height_m/s',
        'Gust_On (0/1)',
        'Wind_Peak_10m_m/s',
        'Wind_Peak_at_Height_m/s',
        'Wind_Pressure_Pa',
        'U_avg_Displacement_mm',
        'S_avg_VonMises_Stress_MPa',
        'Strain_avg',
        'Run_Status'
    ]
    write_xlsx(out_xlsx, run_tag, headers, results_rows)
    print('%s results written to: %s' % (run_tag, os.path.abspath(out_xlsx)))
    return results_rows

no_gust_xlsx = os.path.join(OUTPUT_DIR, 'tapered_no_gust_100.xlsx')
all_gust_xlsx = os.path.join(OUTPUT_DIR, 'tapered_all_gust_100.xlsx')
summary_xlsx = os.path.join(OUTPUT_DIR, 'tapered_height_summary_100.xlsx')
case_summary_xlsx = os.path.join(OUTPUT_DIR, 'tapered_case_summary_100.xlsx')

case_summary_rows = []

no_gust_rows = run_batch('NO_GUST', v_mean_samples, v_peak_no, gust_flags_no, no_gust_xlsx, case_summary_rows)
all_gust_rows = run_batch('ALL_GUST', v_mean_samples, v_peak_all, gust_flags_all, all_gust_xlsx, case_summary_rows)

summary_rows = []
for y_target in height_levels_mm:
    height_m = y_target / 1000.0

    no_rows_h = [r for r in no_gust_rows if abs(r[1] - height_m) < 1.0e-9 and r[11] == 'COMPLETED']
    gust_rows_h = [r for r in all_gust_rows if abs(r[1] - height_m) < 1.0e-9 and r[11] == 'COMPLETED']

    mean_norm_wind = (sum([r[3] for r in no_rows_h]) / float(len(no_rows_h))) if len(no_rows_h) > 0 else None
    mean_gust_wind = (sum([r[6] for r in gust_rows_h]) / float(len(gust_rows_h))) if len(gust_rows_h) > 0 else None
    mean_u_no = (sum([r[8] for r in no_rows_h]) / float(len(no_rows_h))) if len(no_rows_h) > 0 else None
    mean_u_gust = (sum([r[8] for r in gust_rows_h]) / float(len(gust_rows_h))) if len(gust_rows_h) > 0 else None
    mean_s_no = (sum([r[9] for r in no_rows_h]) / float(len(no_rows_h))) if len(no_rows_h) > 0 else None
    mean_s_gust = (sum([r[9] for r in gust_rows_h]) / float(len(gust_rows_h))) if len(gust_rows_h) > 0 else None
    mean_e_no = (sum([r[10] for r in no_rows_h]) / float(len(no_rows_h))) if len(no_rows_h) > 0 else None
    mean_e_gust = (sum([r[10] for r in gust_rows_h]) / float(len(gust_rows_h))) if len(gust_rows_h) > 0 else None

    summary_rows.append([
        height_m,
        mean_norm_wind,
        mean_gust_wind,
        mean_u_no,
        mean_u_gust,
        mean_s_no,
        mean_s_gust,
        mean_e_no,
        mean_e_gust
    ])

summary_headers = [
    'Height_m',
    'Mean_Normal_Wind_at_Height_m/s',
    'Mean_Gust_Wind_at_Height_m/s',
    'Mean_U_NoGust_mm',
    'Mean_U_Gust_mm',
    'Mean_S_NoGust_MPa',
    'Mean_S_Gust_MPa',
    'Mean_Strain_NoGust',
    'Mean_Strain_Gust'
]
write_xlsx(summary_xlsx, 'Height_Summary', summary_headers, summary_rows)

case_summary_headers = [
    'Run_Tag',
    'Case_ID',
    'Wind_Mean_10m_m/s',
    'Wind_Peak_10m_m/s',
    'Gust_On (0/1)',
    'Run_Status',
    'Max_U_mm',
    'Max_S_MPa',
    'Max_Strain'
]
write_xlsx(case_summary_xlsx, 'Case_Summary', case_summary_headers, case_summary_rows)

print('NO_GUST results: %s' % os.path.abspath(no_gust_xlsx))
print('ALL_GUST results: %s' % os.path.abspath(all_gust_xlsx))
print('HEIGHT SUMMARY results: %s' % os.path.abspath(summary_xlsx))
print('CASE SUMMARY results: %s' % os.path.abspath(case_summary_xlsx))

