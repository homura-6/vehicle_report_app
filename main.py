from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
import uvicorn
from datetime import datetime
import pandas as pd
import os
import json

app = FastAPI()

EXCEL_PATH = "Database_Model_X_Price.xlsx"
# 改用 JSON 文件存储完整 Case 数据
CASES_JSON_PATH = "Database_Saved_Cases.json"
CASES_DB_PATH = "Database_Saved_Cases.xlsx" # 保留 Excel 供你另作用途

# 辅助函数：读取和写入 JSON 数据库
def load_json_db():
    if not os.path.exists(CASES_JSON_PATH):
        return {}
    try:
        with open(CASES_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_json_db(data_dict):
    with open(CASES_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data_dict, f, ensure_ascii=False, indent=2)

def parse_and_format_date(date_str: str) -> str:
    if not date_str:
        return ""
    for fmt in ("%Y-%m-%d", "%d %b %Y"):
        try:
            dt = datetime.strptime(date_str.strip(), fmt)
            return dt.strftime("%d %b %Y")
        except ValueError:
            pass
    return date_str

# --- API：生成 Job Number ---
@app.get("/api/generate_job_no")
async def generate_job_no(inspection_date: str):
    try:
        date_obj = None
        for fmt in ("%Y-%m-%d", "%d %b %Y"):
            try:
                date_obj = datetime.strptime(inspection_date.strip(), fmt)
                break
            except ValueError:
                pass

        if not date_obj:
            return {"job_no": "ERROR: Invalid Date"}

        year = date_obj.strftime("%Y")
        month = date_obj.strftime("%m")
        base_calc_val = date_obj.day * 12

        db = load_json_db()
        existing_jobs = list(db.keys())

        while True:
            day_calc_str = str(base_calc_val).zfill(3)
            job_no = f"TC/6/{year}/{month}{day_calc_str}"
            if job_no in existing_jobs:
                base_calc_val += 1
            else:
                break

        return {"job_no": job_no}
    except Exception:
        return {"job_no": "ERROR: Invalid Date"}

# --- API：获取所有 Cases 列表 ---
@app.get("/api/get_cases")
async def get_cases():
    db = load_json_db()
    cases_list = []
    for job_no, data in db.items():
        cases_list.append({
            "Job Number": job_no,
            "Batch No.": data.get("batchNo", ""),
            "Vehicle No.": data.get("vehicleNo", ""),
            "Make & Model": data.get("makeModel", ""),
            "Date Of Accident": data.get("accidentDate", ""),
            "Total Repair Cost": data.get("totalAss", 0)
        })
    return {"cases": cases_list}

# --- API：获取单个 Case 完整数据 ---
@app.get("/api/get_case_detail")
async def get_case_detail(job_no: str):
    db = load_json_db()
    if job_no not in db:
        raise HTTPException(status_code=404, detail="Case not found")
    return {"status": "success", "data": db[job_no]}

# --- API：删除指定 Case ---
@app.delete("/api/delete_case")
async def delete_case(job_no: str):
    db = load_json_db()
    if job_no in db:
        del db[job_no]
        save_json_db(db)
        return {"status": "success", "message": f"Case {job_no} 已彻底删除！"}
    return {"status": "error", "message": "Job Number未找到"}

# --- API：保存完整 Case 数据 ---
@app.post("/api/save_case")
async def save_case(payload: dict):
    try:
        job_no = payload.get("job_no")
        if not job_no:
            return {"status": "error", "message": "Job Number 不能为空"}

        # 1. 全量保存所有表单数据到 JSON
        db = load_json_db()
        db[job_no] = payload
        save_json_db(db)

        # 2. 同步导出简易 Excel 供你另作用途
        record = {
            "Batch No.": payload.get("batchNo", ""),
            "Vehicle No.": payload.get("vehicleNo", ""),
            "Job Number": job_no,
            "Make & Model": payload.get("makeModel", ""),
            "Date Of Accident": payload.get("accidentDate", ""),
            "Date Of Inspection": payload.get("inspectionDate", ""),
            "Report Finish Date": payload.get("finishDate", ""),
            "Total Repair Cost": payload.get("totalAss", 0),
            "Repair Days": payload.get("repairDays", "")
        }
        new_row_df = pd.DataFrame([record])
        if os.path.exists(CASES_DB_PATH):
            try:
                cases_df = pd.read_excel(CASES_DB_PATH)
                cases_df = cases_df[cases_df['Job Number'].astype(str) != job_no]
                cases_df = pd.concat([cases_df, new_row_df], ignore_index=True)
                cases_df.to_excel(CASES_DB_PATH, index=False)
            except Exception:
                new_row_df.to_excel(CASES_DB_PATH, index=False)
        else:
            new_row_df.to_excel(CASES_DB_PATH, index=False)

        return {"status": "success", "message": f"Case {job_no} 已成功全量保存！"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

# --- 其他查询 API ---
@app.get("/api/get_model_info")
async def get_model_info(model_code: str):
    if not os.path.exists(EXCEL_PATH):
        return {"description": "Excel not found", "default_percentage": 100}
    try:
        df = pd.read_excel(EXCEL_PATH)
        df['Model Code Clean'] = df['Model Code'].astype(str).str.strip().str.upper()
        matched = df[df['Model Code Clean'] == model_code.strip().upper()]
        if matched.empty:
            return {"description": "Model Not Found", "default_percentage": 100}
        desc = str(matched.iloc[0].get('Model Description', ''))
        pct = float(matched.iloc[0].get('Default Percentage', 150))
        return {"description": desc, "default_percentage": pct}
    except Exception as e:
        return {"description": str(e), "default_percentage": 100}

@app.get("/api/get_part_price")
async def get_part_price(model_code: str, part_name: str):
    if not os.path.exists(EXCEL_PATH):
        return {"price": 0.0}
    try:
        df = pd.read_excel(EXCEL_PATH)
        df['Model Code Clean'] = df['Model Code'].astype(str).str.strip().str.upper()
        df['Parts Name Clean'] = df['Parts Name'].astype(str).str.strip().str.upper()
        matched = df[(df['Model Code Clean'] == model_code.strip().upper()) & (df['Parts Name Clean'] == part_name.strip().upper())]
        if not matched.empty:
            return {"price": float(matched.iloc[0].get('Price', 0.0))}
        return {"price": 0.0}
    except Exception:
        return {"price": 0.0}

@app.get("/", response_class=HTMLResponse)
async def get_form():
    html_content = """
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Vehicle Assessment Form</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-slate-100 p-4 md:p-8 font-sans">
        <div class="max-w-5xl mx-auto bg-slate-800 text-white rounded-t-xl p-3 flex justify-between items-center shadow-md sticky top-0 z-50">
            <div class="flex space-x-2">
                <button onclick="confirmExit()" class="bg-slate-600 hover:bg-slate-500 px-3 py-1 rounded text-xs">Menu</button>
                <button onclick="openCaseListModal()" class="bg-purple-600 hover:bg-purple-500 px-3 py-1 rounded text-xs font-bold">📂 Open Cases</button>
                <button id="backBtn" onclick="goBackToStep1()" class="hidden bg-slate-600 hover:bg-slate-500 px-3 py-1 rounded text-xs">&larr; Back</button>
            </div>
            <div class="text-xs md:text-sm font-semibold" id="topBarInfo">NEW CASE - Basic Info</div>
            <div>
                <button id="topEditBtn" onclick="enableEditStep1()" class="hidden bg-amber-500 hover:bg-amber-400 px-3 py-1 rounded text-xs font-bold mr-2">Edit Basic Info</button>
                <button id="topSaveBtn" onclick="saveBasicInfo()" class="bg-blue-500 hover:bg-blue-400 px-3 py-1 rounded text-xs font-bold shadow">Save Basic Info</button>
                <button id="topSaveStep2Btn" onclick="lockAndSaveCase()" class="hidden bg-green-500 hover:bg-green-400 px-3 py-1 rounded text-xs font-bold shadow">🔒 Save & Lock Case</button>
                <button id="topEditStep2Btn" onclick="unlockStep2()" class="hidden bg-amber-500 hover:bg-amber-400 px-3 py-1 rounded text-xs font-bold">✏️ Edit Details</button>
            </div>
        </div>

        <!-- Step 1 Page -->
        <div id="step1Page" class="max-w-5xl mx-auto bg-white rounded-b-xl shadow-md p-6 md:p-8 mb-10">
            <h1 class="text-2xl font-bold text-slate-800 mb-6 border-b pb-3">🚗 Step 1: Basic Vehicle Info</h1>
            <form id="basicForm" class="space-y-6">
                <div class="bg-blue-50 p-4 rounded border border-blue-100 mb-6 flex flex-col md:flex-row gap-4">
                    <div class="flex-1">
                        <label class="block text-xs font-bold text-blue-800 uppercase mb-1">Batch Number</label>
                        <input type="number" id="batchNo" class="w-full border rounded p-2 text-sm outline-none step1-input" placeholder="Enter Batch No...">
                    </div>
                    <div class="flex-1">
                        <label class="block text-xs font-bold text-blue-800 uppercase mb-1">Job Number</label>
                        <input type="text" id="jobNo" readonly class="w-full border border-blue-300 bg-blue-100 rounded p-2 text-sm font-bold text-blue-900 outline-none" placeholder="Click SAVE to generate">
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">1. Owner Name</label>
                        <input type="text" id="ownerNameInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div class="md:row-span-2">
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">2. Owner Address</label>
                        <div class="space-y-2">
                            <input type="text" id="ownerAddr1" placeholder="Line 1" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="ownerAddr2" placeholder="Line 2" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="ownerAddr3" placeholder="Line 3" class="w-full border rounded p-2 text-sm step1-input">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">3. Vehicle No.</label>
                        <input type="text" id="vehicleNo" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">4. Make & Model</label>
                        <input type="text" id="makeModelInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">5. Year Make</label>
                        <input type="number" id="yearMakeInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">6. Chassis No.</label>
                        <input type="text" id="chassisNoInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">7. Engine No.</label>
                        <input type="text" id="engineNoInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">8. Engine Capacity (cc)</label>
                        <input type="number" id="engineCapInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">9. Power Rating (kW)</label>
                        <input type="number" id="powerRatingInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">10. Motor No.</label>
                        <input type="text" id="motorNoInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">11. Colour</label>
                        <input type="text" id="colourInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">12. Mileage (km)</label>
                        <input type="number" id="mileageInput" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                </div>

                <div class="border-t pt-4 bg-slate-50 p-4 rounded-lg border">
                    <label class="block text-sm font-bold text-slate-700 mb-2">13. TYRES CONDITION</label>
                    <div class="mb-3">
                        <select id="wheelCount" onchange="renderTyreTable()" class="border rounded p-2 text-sm bg-white outline-none step1-input">
                            <option value="">-- Please select the number of tires --</option>
                            <option value="2">2 Wheel (Motorcycle)</option>
                            <option value="4">4 Wheel (Car / Van)</option>
                            <option value="6">6 Wheel (Truck)</option>
                            <option value="8">8 Wheel (Heavy Truck)</option>
                            <option value="10">10 Wheel (Trailer)</option>
                        </select>
                    </div>
                    <div id="tyreTableContainer">
                        <p class="text-xs text-gray-400 italic">The table will automatically expand after you select the number of tires</p>
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-4 border-t pt-4">
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">14. Damaged Position</label>
                        <textarea id="damagedPosInput" rows="2" class="w-full border rounded p-2 text-sm step1-input"></textarea>
                    </div>
                    <div class="md:row-span-3">
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">18. Workshop Name & Address</label>
                        <div class="space-y-2">
                            <input type="text" id="wsName" placeholder="Workshop Name" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="wsAddr1" placeholder="Address Line 1" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="wsAddr2" placeholder="Address Line 2" class="w-full border rounded p-2 text-sm step1-input">
                            <input type="text" id="wsAddr3" placeholder="Address Line 3" class="w-full border rounded p-2 text-sm step1-input">
                        </div>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">15. Date of Accident</label>
                        <input type="text" id="accidentDate" placeholder="e.g. 12 Oct 2026" onfocus="(this.type='date')" onblur="handleDateBlur(this)" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">16. Date of Inspection</label>
                        <input type="text" id="inspectionDate" placeholder="e.g. 12 Oct 2026" onfocus="(this.type='date')" onchange="autoCalculateFinishDate()" onblur="handleDateBlur(this)" class="w-full border rounded p-2 text-sm step1-input">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-gray-600 uppercase mb-1">17. Report Finish Date</label>
                        <input type="text" id="finishDate" readonly placeholder="Auto-calculated" class="w-full border rounded p-2 text-sm bg-gray-100 text-gray-600">
                    </div>
                </div>

                <div class="pt-6 flex justify-between items-center border-t">
                    <button type="button" id="mainSaveBtn" onclick="saveBasicInfo()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-8 rounded-lg shadow-md">
                        💾 SAVE & GENERATE JOB NO.
                    </button>
                    <button type="button" id="nextBtn" onclick="goToStep2()" class="hidden bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-8 rounded-lg shadow-md">
                        NEXT: Damage Details &rarr;
                    </button>
                </div>
            </form>
        </div>

        <!-- Step 2 Page -->
        <div id="step2Page" class="hidden max-w-5xl mx-auto bg-white rounded-b-xl shadow-md p-6 md:p-8 mb-10">
            <h1 class="text-2xl font-bold text-slate-800 mb-6 border-b pb-3">🛠️ Step 2: Assessment Details & Cost</h1>

            <!-- PARTS -->
            <div class="mb-10 bg-slate-50 p-4 rounded-xl border border-slate-200">
                <div class="flex flex-col mb-4 gap-3">
                    <h2 class="text-lg font-bold text-slate-800">PARTS ASSESSMENT</h2>
                    <div class="bg-white p-3 rounded-lg border shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                        <div class="flex items-center space-x-3">
                            <div class="flex items-center space-x-1">
                                <label class="text-xs font-bold text-slate-600">MODEL CODE:</label>
                                <input type="text" id="modelCodeInput" placeholder="e.g. A5" onblur="fetchModelInfo()" class="border rounded px-2 py-1 text-xs uppercase focus:ring-2 focus:ring-blue-500 outline-none step2-input w-28 font-bold">
                            </div>
                            <div class="flex items-center space-x-1">
                                <label class="text-xs font-bold text-slate-600">[%]:</label>
                                <input type="number" id="globalPercentage" value="150" step="1" onchange="recalcAllPartsRows()" onkeyup="recalcAllPartsRows()" class="border rounded px-2 py-1 text-xs w-20 text-center font-bold text-blue-600 step2-input">
                                <span class="text-xs font-bold text-slate-600">%</span>
                            </div>
                        </div>
                        <div class="text-xs text-slate-600 font-medium bg-slate-50 p-2 rounded border flex-1 md:text-right" id="modelDescriptionText">
                            Description: <span class="italic text-gray-400">Please enter model code</span>
                        </div>
                    </div>
                </div>

                <div class="overflow-x-auto">
                    <table class="w-full text-left bg-white rounded border text-xs">
                        <thead>
                            <tr class="bg-slate-200 text-slate-700 font-bold uppercase">
                                <th class="p-2 border w-10 text-center">S/N</th>
                                <th class="p-2 border w-14">QTY</th>
                                <th class="p-2 border">PARTS NAME</th>
                                <th class="p-2 border w-24">CODE</th>
                                <th class="p-2 border w-28">CONDITION</th>
                                <th class="p-2 border w-28">NETT PRICE (S$)</th>
                                <th class="p-2 border w-28">ESTIMATE (S$)</th>
                                <th class="p-2 border w-28">ASSESSMENT (S$)</th>
                                <th class="p-2 border w-10 text-center">DEL</th>
                            </tr>
                        </thead>
                        <tbody id="partsTbody"></tbody>
                    </table>
                </div>
                <button onclick="addPartsRow()" class="step2-btn mt-3 bg-blue-600 hover:bg-blue-700 text-white font-bold px-3 py-1.5 rounded text-xs">
                    + Add Part Row
                </button>
                <div class="mt-4 border-t pt-3 text-xs flex justify-end">
                    <div class="w-80 space-y-1 bg-white p-3 rounded border shadow-sm">
                        <div class="flex justify-between"><span class="text-gray-600">Sub Total:</span><span class="font-bold">S$ <span id="partsSubEst">0.00</span> | S$ <span id="partsSubAss">0.00</span></span></div>
                        <div class="flex justify-between text-red-600"><span>Less: Discount 10%:</span><span class="font-bold">-S$ <span id="partsDiscEst">0.00</span> | -S$ <span id="partsDiscAss">0.00</span></span></div>
                        <div class="flex justify-between text-sm font-bold border-t pt-1 text-blue-900"><span>Total Parts:</span><span>S$ <span id="partsTotalEst">0.00</span> | S$ <span id="partsTotalAss">0.00</span></span></div>
                    </div>
                </div>
            </div>

            <!-- MISCELLANEOUS -->
            <div class="mb-10 bg-slate-50 p-4 rounded-xl border border-slate-200">
                <h2 class="text-lg font-bold text-slate-800 mb-3">MISCELLANEOUS ITEMS</h2>
                <div class="overflow-x-auto">
                    <table class="w-full text-left bg-white rounded border text-xs">
                        <thead>
                            <tr class="bg-slate-200 text-slate-700 font-bold uppercase">
                                <th class="p-2 border w-10 text-center">S/N</th>
                                <th class="p-2 border w-14">QTY</th>
                                <th class="p-2 border">DESCRIPTION</th>
                                <th class="p-2 border w-24">CODE</th>
                                <th class="p-2 border w-36">CONDITION</th>
                                <th class="p-2 border w-28">ESTIMATE (S$)</th>
                                <th class="p-2 border w-28">ASSESSMENT (S$)</th>
                                <th class="p-2 border w-10 text-center">DEL</th>
                            </tr>
                        </thead>
                        <tbody id="miscTbody"></tbody>
                    </table>
                </div>
                <div class="flex justify-between items-center mt-3">
                    <button onclick="addMiscRow()" class="step2-btn bg-blue-600 hover:bg-blue-700 text-white font-bold px-3 py-1.5 rounded text-xs">+ Add Misc Row</button>
                    <div class="text-xs bg-white p-2 rounded border font-bold text-slate-800">Misc Total: S$ <span id="miscTotalEst">0.00</span> | S$ <span id="miscTotalAss">0.00</span></div>
                </div>
            </div>

            <!-- LABOUR -->
            <div class="mb-10 bg-slate-50 p-4 rounded-xl border border-slate-200">
                <h2 class="text-lg font-bold text-slate-800 mb-3">PANEL & MECHANICAL LABOUR</h2>
                <div id="labourChecklist" class="space-y-2 mb-4"></div>
                <div class="border-t pt-3 mt-3">
                    <div class="text-xs font-bold text-slate-600 mb-2 uppercase">Custom Labour Items</div>
                    <div id="customLabourContainer" class="space-y-2"></div>
                    <button onclick="addCustomLabourRow()" class="step2-btn mt-2 bg-slate-700 hover:bg-slate-800 text-white font-bold px-3 py-1.5 rounded text-xs">+ Add Custom Labour Row</button>
                </div>
                <div class="flex justify-end mt-3 text-xs bg-white p-2 rounded border font-bold text-slate-800">Labour Total: S$ <span id="labourTotalEst">0.00</span> | S$ <span id="labourTotalAss">0.00</span></div>
            </div>

            <!-- PAINT -->
            <div class="mb-10 bg-slate-50 p-4 rounded-xl border border-slate-200">
                <h2 class="text-lg font-bold text-slate-800 mb-3">PAINT</h2>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                        <h3 class="font-bold text-xs text-slate-600 uppercase mb-2">Paint Labour Cost</h3>
                        <div class="flex gap-2">
                            <input type="text" value="Paint Labour" readonly class="flex-1 border rounded p-1.5 text-xs bg-gray-50">
                            <input type="text" placeholder="Est (S$)" id="paintLabourEst" oninput="formatNumberInput(this)" onkeyup="calcGrandTotals()" class="w-28 border rounded p-1.5 text-xs font-mono step2-input">
                            <input type="text" placeholder="Ass (S$)" id="paintLabourAss" oninput="formatNumberInput(this)" onkeyup="calcGrandTotals()" class="w-28 border rounded p-1.5 text-xs font-mono font-bold step2-input">
                        </div>
                    </div>
                    <div>
                        <h3 class="font-bold text-xs text-slate-600 uppercase mb-2">Paint Material Cost</h3>
                        <div class="flex gap-2">
                            <input type="text" value="Paint Material" readonly class="flex-1 border rounded p-1.5 text-xs bg-gray-50">
                            <input type="text" placeholder="Est (S$)" id="paintMatEst" oninput="formatNumberInput(this)" onkeyup="calcGrandTotals()" class="w-28 border rounded p-1.5 text-xs font-mono step2-input">
                            <input type="text" placeholder="Ass (S$)" id="paintMatAss" oninput="formatNumberInput(this)" onkeyup="calcGrandTotals()" class="w-28 border rounded p-1.5 text-xs font-mono font-bold step2-input">
                        </div>
                    </div>
                </div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-4 bg-blue-50 p-4 rounded-xl border border-blue-100 mb-10">
                <div>
                    <label class="block text-xs font-bold text-blue-900 uppercase mb-1">19. Working Days (Repair Days)</label>
                    <input type="text" id="repairDaysInput" placeholder="e.g. 5" class="w-full border rounded p-2 text-sm step2-input">
                </div>
                <div>
                    <label class="block text-xs font-bold text-blue-900 uppercase mb-1">20. Photo No. (Quantity)</label>
                    <input type="text" id="photoQtyInput" placeholder="e.g. 12" class="w-full border rounded p-2 text-sm step2-input">
                </div>
            </div>

            <div class="flex justify-end gap-4 mb-8">
                <button id="bottomSaveStep2Btn" onclick="lockAndSaveCase()" class="bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-8 rounded-lg shadow-md">🔒 LOCK & SAVE CASE TO DB</button>
                <button id="bottomEditStep2Btn" onclick="unlockStep2()" class="hidden bg-amber-500 hover:bg-amber-600 text-white font-bold py-3 px-8 rounded-lg shadow-md">✏️ UNLOCK TO EDIT DETAILS</button>
            </div>

            <!-- FINAL CALCULATION -->
            <div class="bg-slate-800 text-white p-6 rounded-xl shadow-lg">
                <h2 class="text-lg font-bold text-yellow-400 mb-4 uppercase">Final Calculation</h2>
                <table class="w-full text-left text-sm">
                    <thead>
                        <tr class="border-b border-slate-700 text-slate-400 text-xs uppercase">
                            <th class="py-2">Description</th>
                            <th class="py-2 text-right">Repairer's Estimate (S$)</th>
                            <th class="py-2 text-right">Our Assessment (S$)</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-700">
                        <tr><td class="py-2">PARTS</td><td class="text-right" id="finalPartsEst">S$ 0.00</td><td class="text-right" id="finalPartsAss">S$ 0.00</td></tr>
                        <tr><td class="py-2">MISCELLANEOUS ITEMS</td><td class="text-right" id="finalMiscEst">S$ 0.00</td><td class="text-right" id="finalMiscAss">S$ 0.00</td></tr>
                        <tr><td class="py-2">PANEL & MECHANICAL LABOUR</td><td class="text-right" id="finalLabourEst">S$ 0.00</td><td class="text-right" id="finalLabourAss">S$ 0.00</td></tr>
                        <tr><td class="py-2">PAINT LABOUR</td><td class="text-right" id="finalPaintLabourEst">S$ 0.00</td><td class="text-right" id="finalPaintLabourAss">S$ 0.00</td></tr>
                        <tr><td class="py-2">PAINT MATERIAL</td><td class="text-right" id="finalPaintMatEst">S$ 0.00</td><td class="text-right" id="finalPaintMatAss">S$ 0.00</td></tr>
                        <tr class="font-bold text-base text-yellow-400 pt-2">
                            <td class="py-3">TOTAL REPAIR COST</td>
                            <td class="text-right py-3" id="finalTotalEst">S$ 0.00</td>
                            <td class="text-right py-3" id="finalTotalAss">S$ 0.00</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Cases 列表弹窗 -->
        <div id="caseModal" class="hidden fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div class="bg-white rounded-xl shadow-2xl w-full max-w-4xl max-h-[80vh] flex flex-col overflow-hidden">
                <div class="bg-slate-800 text-white p-4 flex justify-between items-center">
                    <h3 class="font-bold text-lg">📂 Saved Cases Database</h3>
                    <button onclick="closeCaseListModal()" class="text-gray-400 hover:text-white font-bold text-xl">✕</button>
                </div>
                <div class="p-4 bg-slate-50 border-b">
                    <input type="text" id="caseSearchInput" onkeyup="filterCasesTable()" placeholder="🔍 Search by Job No, Vehicle No, Batch No..." class="w-full border rounded-lg p-2 text-sm outline-none">
                </div>
                <div class="overflow-y-auto p-4 flex-1">
                    <table class="w-full text-left border text-xs">
                        <thead>
                            <tr class="bg-slate-200 text-slate-700 font-bold uppercase sticky top-0">
                                <th class="p-2 border">Job Number</th>
                                <th class="p-2 border">Batch No.</th>
                                <th class="p-2 border">Vehicle No.</th>
                                <th class="p-2 border">Make & Model</th>
                                <th class="p-2 border">Date Of Accident</th>
                                <th class="p-2 border text-right">Total Assessment</th>
                                <th class="p-2 border text-center">Actions</th>
                            </tr>
                        </thead>
                        <tbody id="caseListTbody"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <script>
            function formatDateToDDMMMYYYY(dateStr) {
                if (!dateStr) return "";
                const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                if (dateStr.includes("-")) {
                    const parts = dateStr.split("-");
                    if (parts.length === 3) {
                        const year = parts[0];
                        const monthIdx = parseInt(parts[1], 10) - 1;
                        const day = String(parseInt(parts[2], 10)).padStart(2, '0');
                        if (!isNaN(monthIdx) && months[monthIdx]) {
                            return `${day} ${months[monthIdx]} ${year}`;
                        }
                    }
                }
                return dateStr;
            }

            function handleDateBlur(inputEl) {
                inputEl.type = 'text';
                if (inputEl.value) inputEl.value = formatDateToDDMMMYYYY(inputEl.value);
            }

            function formatNumber(num) {
                if (isNaN(num) || num === null) return "0.00";
                return Number(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            }

            function parseNumber(str) {
                if (!str) return 0;
                const cleanStr = String(str).replace(/,/g, '');
                return parseFloat(cleanStr) || 0;
            }

            function formatNumberInput(inputEl) {
                let cursorPosition = inputEl.selectionStart;
                let originalLength = inputEl.value.length;
                let val = parseNumber(inputEl.value);
                if (inputEl.value === "" || inputEl.value === "-") return;
                if (inputEl.value.endsWith('.')) return;
                inputEl.value = val.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
                let newLength = inputEl.value.length;
                inputEl.setSelectionRange(cursorPosition + (newLength - originalLength), cursorPosition + (newLength - originalLength));
            }

            const remarkRules = {
                "rp": { wording: "REPAIR", type: "zero" },
                "sv": { wording: "SERVICE", type: "zero" },
                "rj": { wording: "REJECTED", type: "zero" },
                "na": { wording: "NA", type: "zero" },
                "be": { wording: "BENT", type: "full" },
                "de": { wording: "DENT", type: "full" },
                "bu": { wording: "BUCKLED", type: "full" },
                "ne": { wording: "NECESSARY", type: "full" },
                "br": { wording: "BROKEN", type: "full" },
                "cr": { wording: "CRACKED", type: "full" },
                "ma": { wording: "MALFUNCTION", type: "full" },
                "df": { wording: "DEFORMED", type: "full" },
                "dp": { wording: "DEPLOYED", type: "full" },
                "ja": { wording: "JAMMED", type: "full" },
                "to": { wording: "TORN", type: "full" },
                "ru": { wording: "REUSE", type: "full" },
                "gl": { wording: "GLAZED", type: "full" },
                "mi": { wording: "MISSING", type: "full" },
                "we": { wording: "WEAK", type: "full" },
                "cu": { wording: "CUT", type: "full" },
                "sc": { wording: "SCRATCHED", type: "full" },
                "1m": { wording: "1PC MALFUNCTION", type: "m_pcs", pcs: 1 },
                "2m": { wording: "2PCS MALFUNCTION", type: "m_pcs", pcs: 2 },
                "3m": { wording: "3PCS MALFUNCTION", type: "m_pcs", pcs: 3 },
                "4m": { wording: "4PCS MALFUNCTION", type: "m_pcs", pcs: 4 },
                "5m": { wording: "5PCS MALFUNCTION", type: "m_pcs", pcs: 5 },
                "c1": { wording: "Cut/10%", type: "cut", percent: 0.90 },
                "c2": { wording: "Cut/20%", type: "cut", percent: 0.80 },
                "c3": { wording: "Cut/30%", type: "cut", percent: 0.70 },
                "c4": { wording: "Cut/40%", type: "cut", percent: 0.60 },
                "c5": { wording: "Cut/50%", type: "cut", percent: 0.50 },
                "c6": { wording: "Cut/60%", type: "cut", percent: 0.40 },
                "c7": { wording: "Cut/70%", type: "cut", percent: 0.30 },
                "c8": { wording: "Cut/80%", type: "cut", percent: 0.20 },
                "c9": { wording: "Cut/90%", type: "cut", percent: 0.10 }
            };

            const tyrePositions = {
                "2": ["Front", "Rear"],
                "4": ["Front RHS", "Front LHS", "Rear RHS", "Rear LHS"],
                "6": ["Front RHS", "Front LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"],
                "8": ["Front RHS", "Front LHS", "Mid RHS", "Mid LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"],
                "10": ["Front RHS", "Front LHS", "Mid Outer RHS", "Mid Inner RHS", "Mid Inner LHS", "Mid Outer LHS", "Rear Outer RHS", "Rear Inner RHS", "Rear Inner LHS", "Rear Outer LHS"]
            };

            let isStep2Locked = false;
            let allCasesData = [];

            async function openCaseListModal() {
                try {
                    const res = await fetch('/api/get_cases');
                    const data = await res.json();
                    allCasesData = data.cases || [];
                    renderCasesTable(allCasesData);
                    document.getElementById('caseModal').classList.remove('hidden');
                } catch (e) {
                    alert("无法获取历史 Cases 数据！");
                }
            }

            function closeCaseListModal() {
                document.getElementById('caseModal').classList.add('hidden');
            }

            function renderCasesTable(cases) {
                const tbody = document.getElementById('caseListTbody');
                if (cases.length === 0) {
                    tbody.innerHTML = `<tr><td colspan="7" class="text-center p-4 text-gray-400">No case records are currently saved.</td></tr>`;
                    return;
                }
                let html = "";
                cases.forEach(c => {
                    html += `
                        <tr class="border-b hover:bg-blue-50">
                            <td class="p-2 border font-bold text-blue-700">${c['Job Number']}</td>
                            <td class="p-2 border">${c['Batch No.']}</td>
                            <td class="p-2 border font-semibold">${c['Vehicle No.']}</td>
                            <td class="p-2 border">${c['Make & Model']}</td>
                            <td class="p-2 border">${c['Date Of Accident']}</td>
                            <td class="p-2 border text-right font-bold">S$ ${formatNumber(c['Total Repair Cost'])}</td>
                            <td class="p-2 border text-center space-x-2">
                                <button onclick="loadCaseForEdit('${c['Job Number']}')" class="bg-blue-600 hover:bg-blue-700 text-white font-bold px-2 py-1 rounded text-xs">✏️ 修改/查看</button>
                                <button onclick="confirmDeleteCase('${c['Job Number']}')" class="bg-red-600 hover:bg-red-700 text-white font-bold px-2 py-1 rounded text-xs">🗑️ 删除</button>
                            </td>
                        </tr>
                    `;
                });
                tbody.innerHTML = html;
            }

            function filterCasesTable() {
                const query = document.getElementById('caseSearchInput').value.toLowerCase();
                const filtered = allCasesData.filter(c => 
                    String(c['Job Number']).toLowerCase().includes(query) ||
                    String(c['Vehicle No.']).toLowerCase().includes(query) ||
                    String(c['Batch No.']).toLowerCase().includes(query) ||
                    String(c['Make & Model']).toLowerCase().includes(query)
                );
                renderCasesTable(filtered);
            }

            async function confirmDeleteCase(jobNo) {
                if (!confirm(`⚠️ 确定要彻底删除 Case [ ${jobNo} ] 吗？`)) return;
                try {
                    const res = await fetch(`/api/delete_case?job_no=${encodeURIComponent(jobNo)}`, { method: 'DELETE' });
                    const result = await res.json();
                    if (result.status === "success") {
                        alert(result.message);
                        openCaseListModal();
                    } else { alert("删除失败: " + result.message); }
                } catch (e) { alert("网络请求失败！"); }
            }

            // --- 收集全量表单数据的打包函数 ---
            function collectAllFormData() {
                const jobNo = document.getElementById('jobNo').value;
                const batchNo = document.getElementById('batchNo').value;
                const vehicleNo = document.getElementById('vehicleNo').value;
                const makeModel = document.getElementById('makeModelInput').value;
                const accidentDate = formatDateToDDMMMYYYY(document.getElementById('accidentDate').value);
                const inspectionDate = formatDateToDDMMMYYYY(document.getElementById('inspectionDate').value);
                const finishDate = document.getElementById('finishDate').value;

                // Step 1 基础信息
                const step1Data = {
                    ownerName: document.getElementById('ownerNameInput').value,
                    ownerAddr1: document.getElementById('ownerAddr1').value,
                    ownerAddr2: document.getElementById('ownerAddr2').value,
                    ownerAddr3: document.getElementById('ownerAddr3').value,
                    yearMake: document.getElementById('yearMakeInput').value,
                    chassisNo: document.getElementById('chassisNoInput').value,
                    engineNo: document.getElementById('engineNoInput').value,
                    engineCap: document.getElementById('engineCapInput').value,
                    powerRating: document.getElementById('powerRatingInput').value,
                    motorNo: document.getElementById('motorNoInput').value,
                    colour: document.getElementById('colourInput').value,
                    mileage: document.getElementById('mileageInput').value,
                    wheelCount: document.getElementById('wheelCount').value,
                    tyres: [],
                    damagedPos: document.getElementById('damagedPosInput').value,
                    wsName: document.getElementById('wsName').value,
                    wsAddr1: document.getElementById('wsAddr1').value,
                    wsAddr2: document.getElementById('wsAddr2').value,
                    wsAddr3: document.getElementById('wsAddr3').value
                };

                // 抓取轮胎列表
                document.querySelectorAll('#tyreTableContainer tr').forEach((tr, idx) => {
                    if (idx === 0) return;
                    const inputs = tr.querySelectorAll('input');
                    if (inputs.length === 3) {
                        step1Data.tyres.push({
                            brand: inputs[0].value,
                            size: inputs[1].value,
                            mm: inputs[2].value
                        });
                    }
                });

                // Parts 列表
                const parts = [];
                document.querySelectorAll('#partsTbody .parts-row').forEach(row => {
                    const rId = row.id.split('_')[1];
                    parts.push({
                        qty: document.getElementById(`partQty_${rId}`).value,
                        name: document.getElementById(`partNameInput_${rId}`).value,
                        code: document.getElementById(`partCode_${rId}`).value,
                        nett: document.getElementById(`partNett_${rId}`).value
                    });
                });

                // Misc 列表
                const misc = [];
                document.querySelectorAll('#miscTbody .misc-row').forEach(row => {
                    const rId = row.id.split('_')[1];
                    misc.push({
                        qty: document.getElementById(`miscQty_${rId}`).value,
                        desc: row.querySelectorAll('input')[1].value,
                        code: document.getElementById(`miscCode_${rId}`).value,
                        est: document.getElementById(`miscEst_${rId}`).value,
                        ass: document.getElementById(`miscAss_${rId}`).value
                    });
                });

                // Labour 勾选列表
                const labourChecks = [];
                labourItems.forEach((_, idx) => {
                    const checkbox = document.querySelector(`#labourRow_${idx} input[type="checkbox"]`);
                    if (checkbox && checkbox.checked) {
                        labourChecks.push({
                            index: idx,
                            est: document.getElementById(`labourEst_${idx}`).value,
                            ass: document.getElementById(`labourAss_${idx}`).value
                        });
                    }
                });

                // Custom Labour 列表
                const customLabour = [];
                document.querySelectorAll('.custom-labour-row').forEach(row => {
                    const cId = row.id.split('_')[1];
                    customLabour.push({
                        desc: row.querySelector('input[type="text"]').value,
                        est: document.getElementById(`customLabourEst_${cId}`).value,
                        ass: document.getElementById(`customLabourAss_${cId}`).value
                    });
                });

                const totalAss = parseNumber(document.getElementById('finalTotalAss').innerText.replace("S$ ", ""));

                return {
                    job_no: jobNo,
                    batchNo: batchNo,
                    vehicleNo: vehicleNo,
                    makeModel: makeModel,
                    accidentDate: accidentDate,
                    inspectionDate: inspectionDate,
                    finishDate: finishDate,
                    totalAss: totalAss,
                    repairDays: document.getElementById('repairDaysInput').value,
                    photoQty: document.getElementById('photoQtyInput').value,
                    modelCode: document.getElementById('modelCodeInput').value,
                    globalPercentage: document.getElementById('globalPercentage').value,
                    paintLabourEst: document.getElementById('paintLabourEst').value,
                    paintLabourAss: document.getElementById('paintLabourAss').value,
                    paintMatEst: document.getElementById('paintMatEst').value,
                    paintMatAss: document.getElementById('paintMatAss').value,
                    step1Data: step1Data,
                    parts: parts,
                    misc: misc,
                    labourChecks: labourChecks,
                    customLabour: customLabour
                };
            }

            // --- 核心：读档还原全量数据 ---
            async function loadCaseForEdit(jobNo) {
                try {
                    const res = await fetch(`/api/get_case_detail?job_no=${encodeURIComponent(jobNo)}`);
                    const result = await res.json();

                    if (result.status === "success") {
                        const d = result.data;

                        // 1. 还原基础信息
                        document.getElementById('jobNo').value = d.job_no || "";
                        document.getElementById('batchNo').value = d.batchNo || "";
                        document.getElementById('vehicleNo').value = d.vehicleNo || "";
                        document.getElementById('makeModelInput').value = d.makeModel || "";
                        document.getElementById('accidentDate').value = d.accidentDate || "";
                        document.getElementById('inspectionDate').value = d.inspectionDate || "";
                        document.getElementById('finishDate').value = d.finishDate || "";

                        const s1 = d.step1Data || {};
                        document.getElementById('ownerNameInput').value = s1.ownerName || "";
                        document.getElementById('ownerAddr1').value = s1.ownerAddr1 || "";
                        document.getElementById('ownerAddr2').value = s1.ownerAddr2 || "";
                        document.getElementById('ownerAddr3').value = s1.ownerAddr3 || "";
                        document.getElementById('yearMakeInput').value = s1.yearMake || "";
                        document.getElementById('chassisNoInput').value = s1.chassisNo || "";
                        document.getElementById('engineNoInput').value = s1.engineNo || "";
                        document.getElementById('engineCapInput').value = s1.engineCap || "";
                        document.getElementById('powerRatingInput').value = s1.powerRating || "";
                        document.getElementById('motorNoInput').value = s1.motorNo || "";
                        document.getElementById('colourInput').value = s1.colour || "";
                        document.getElementById('mileageInput').value = s1.mileage || "";
                        document.getElementById('damagedPosInput').value = s1.damagedPos || "";
                        document.getElementById('wsName').value = s1.wsName || "";
                        document.getElementById('wsAddr1').value = s1.wsAddr1 || "";
                        document.getElementById('wsAddr2').value = s1.wsAddr2 || "";
                        document.getElementById('wsAddr3').value = s1.wsAddr3 || "";

                        // 轮胎还原
                        if (s1.wheelCount) {
                            document.getElementById('wheelCount').value = s1.wheelCount;
                            renderTyreTable();
                            if (s1.tyres) {
                                document.querySelectorAll('#tyreTableContainer tr').forEach((tr, idx) => {
                                    if (idx > 0 && s1.tyres[idx - 1]) {
                                        const inputs = tr.querySelectorAll('input');
                                        inputs[0].value = s1.tyres[idx - 1].brand || "";
                                        inputs[1].value = s1.tyres[idx - 1].size || "";
                                        inputs[2].value = s1.tyres[idx - 1].mm || "";
                                    }
                                });
                            }
                        }

                        // 2. 还原 Step 2 配置
                        document.getElementById('modelCodeInput').value = d.modelCode || "";
                        document.getElementById('globalPercentage').value = d.globalPercentage || 150;
                        document.getElementById('repairDaysInput').value = d.repairDays || "";
                        document.getElementById('photoQtyInput').value = d.photoQty || "";
                        document.getElementById('paintLabourEst').value = d.paintLabourEst || "";
                        document.getElementById('paintLabourAss').value = d.paintLabourAss || "";
                        document.getElementById('paintMatEst').value = d.paintMatEst || "";
                        document.getElementById('paintMatAss').value = d.paintMatAss || "";

                        // 还原 Parts 表格
                        document.getElementById('partsTbody').innerHTML = "";
                        partsRowCounter = 0;
                        if (d.parts && d.parts.length > 0) {
                            d.parts.forEach(p => {
                                addPartsRow();
                                const rId = partsRowCounter;
                                document.getElementById(`partQty_${rId}`).value = p.qty;
                                document.getElementById(`partNameInput_${rId}`).value = p.name;
                                document.getElementById(`partCode_${rId}`).value = p.code;
                                document.getElementById(`partNett_${rId}`).value = p.nett;
                                calcPartsRow(rId);
                            });
                        }

                        // 还原 Misc 表格
                        document.getElementById('miscTbody').innerHTML = "";
                        miscRowCounter = 0;
                        if (d.misc && d.misc.length > 0) {
                            d.misc.forEach(m => {
                                addMiscRow();
                                const rId = miscRowCounter;
                                document.getElementById(`miscQty_${rId}`).value = m.qty;
                                document.querySelectorAll(`#miscRow_${rId} input`)[1].value = m.desc;
                                document.getElementById(`miscCode_${rId}`).value = m.code;
                                document.getElementById(`miscEst_${rId}`).value = m.est;
                                document.getElementById(`miscAss_${rId}`).value = m.ass;
                                updateMiscCondition(rId);
                            });
                        }

                        // 还原 Labour 勾选
                        labourItems.forEach((_, idx) => {
                            const cb = document.querySelector(`#labourRow_${idx} input[type="checkbox"]`);
                            if (cb && cb.checked) cb.click();
                        });
                        if (d.labourChecks) {
                            d.labourChecks.forEach(lc => {
                                const cb = document.querySelector(`#labourRow_${lc.index} input[type="checkbox"]`);
                                if (cb && !cb.checked) {
                                    cb.click();
                                    document.getElementById(`labourEst_${lc.index}`).value = lc.est;
                                    document.getElementById(`labourAss_${lc.index}`).value = lc.ass;
                                }
                            });
                        }

                        // 还原 Custom Labour
                        document.getElementById('customLabourContainer').innerHTML = "";
                        customLabourCounter = 0;
                        if (d.customLabour) {
                            d.customLabour.forEach(cl => {
                                addCustomLabourRow();
                                const cId = customLabourCounter;
                                document.querySelector(`#customLabourRow_${cId} input[type="text"]`).value = cl.desc;
                                document.getElementById(`customLabourEst_${cId}`).value = cl.est;
                                document.getElementById(`customLabourAss_${cId}`).value = cl.ass;
                            });
                        }

                        fetchModelInfo();
                        calcPartsTotals();
                        calcMiscTotals();
                        calcLabourTotals();

                        document.getElementById('topBarInfo').innerHTML = 
                            `<span class="text-gray-400">Batch:</span> ${d.batchNo || 'N/A'} | 
                             <span class="text-gray-400">Veh:</span> ${d.vehicleNo || 'N/A'} | 
                             <span class="text-yellow-400 font-bold">${d.job_no}</span>`;

                        enableEditStep1();
                        unlockStep2();
                        closeCaseListModal();
                        alert(`已完全载入 Case: ${jobNo}！全量资料已还原。`);
                    }
                } catch (e) {
                    alert("载入 Case 失败！");
                }
            }

            // --- 保存函数 ---
            async function lockAndSaveCase() {
                const jobNo = document.getElementById('jobNo').value;
                if (!jobNo || jobNo.startsWith("ERROR")) {
                    alert("错误：无效的 Job Number，无法存档！");
                    return;
                }

                const payload = collectAllFormData();

                try {
                    const response = await fetch('/api/save_case', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    });
                    const result = await response.json();

                    if (result.status === "success") {
                        isStep2Locked = true;
                        document.querySelectorAll('#step2Page input, #step2Page select, #step2Page textarea').forEach(el => el.setAttribute('disabled', 'true'));
                        document.querySelectorAll('.step2-btn, .del-btn').forEach(btn => btn.classList.add('hidden'));

                        document.getElementById('topSaveStep2Btn').classList.add('hidden');
                        document.getElementById('topEditStep2Btn').classList.remove('hidden');
                        document.getElementById('bottomSaveStep2Btn').classList.add('hidden');
                        document.getElementById('bottomEditStep2Btn').classList.remove('hidden');

                        alert("🔒 " + result.message);
                    } else {
                        alert("存档失败: " + result.message);
                    }
                } catch (e) {
                    alert("网络请求失败，无法保存。");
                }
            }

            async function saveBasicInfo() {
                const inspectionDate = document.getElementById('inspectionDate').value;
                const batchNo = document.getElementById('batchNo').value;
                const vehicleNo = document.getElementById('vehicleNo').value;
                let currentJobNo = document.getElementById('jobNo').value;

                if (!currentJobNo || currentJobNo.startsWith("ERROR")) {
                    if (!inspectionDate) {
                        alert("请先选择 16. Date of Inspection，才能自动计算生成 Job Number!");
                        return;
                    }
                    const res = await fetch(`/api/generate_job_no?inspection_date=${encodeURIComponent(inspectionDate)}`);
                    const data = await res.json();
                    currentJobNo = data.job_no;
                    document.getElementById('jobNo').value = currentJobNo;
                }

                document.getElementById('topBarInfo').innerHTML = 
                    `<span class="text-gray-400">Batch:</span> ${batchNo || 'N/A'} | 
                     <span class="text-gray-400">Veh:</span> ${vehicleNo || 'N/A'} | 
                     <span class="text-yellow-400 font-bold">${currentJobNo}</span>`;

                document.querySelectorAll('.step1-input').forEach(i => i.setAttribute('disabled', true));
                document.getElementById('mainSaveBtn').classList.add('hidden');
                document.getElementById('topSaveBtn').classList.add('hidden');
                document.getElementById('topEditBtn').classList.remove('hidden');
                document.getElementById('nextBtn').classList.remove('hidden');

                // 产生 Job No 后立刻做一次全量存档
                await lockAndSaveCase();
            }

            // --- 剩余表单交互辅助逻辑 ---
            async function fetchModelInfo() {
                const modelCode = document.getElementById('modelCodeInput').value.trim();
                if (!modelCode) return;
                try {
                    const res = await fetch(`/api/get_model_info?model_code=${encodeURIComponent(modelCode)}`);
                    const data = await res.json();
                    document.getElementById('modelDescriptionText').innerHTML = `Description: <span class="font-bold text-slate-800">${data.description}</span>`;
                    document.getElementById('globalPercentage').value = data.default_percentage;
                    recalcAllPartsRows();
                } catch (e) {}
            }

            async function fetchPartPriceByInput(rowId) {
                const modelCode = document.getElementById('modelCodeInput').value.trim();
                const partNameInput = document.getElementById(`partNameInput_${rowId}`);
                const partName = partNameInput ? partNameInput.value.trim() : "";
                const nettInput = document.getElementById(`partNett_${rowId}`);
                if (!modelCode || !partName) return;
                try {
                    const res = await fetch(`/api/get_part_price?model_code=${encodeURIComponent(modelCode)}&part_name=${encodeURIComponent(partName)}`);
                    const data = await res.json();
                    if (data.price > 0) nettInput.value = formatNumber(data.price);
                    calcPartsRow(rowId);
                } catch (e) {}
            }

            function renderTyreTable() {
                const count = document.getElementById('wheelCount').value;
                const container = document.getElementById('tyreTableContainer');
                if (!count || !tyrePositions[count]) {
                    container.innerHTML = `<p class="text-xs text-gray-400 italic">The table will automatically expand after you select the number of tires</p>`;
                    return;
                }
                let html = `<table class="w-full text-left border bg-white text-xs mt-2">
                    <tr class="bg-slate-200 font-bold">
                        <th class="p-2 border">Tyres Position</th><th class="p-2 border">Make</th><th class="p-2 border">Size</th><th class="p-2 border">Balance (mm)</th>
                    </tr>`;
                tyrePositions[count].forEach(pos => {
                    html += `<tr>
                        <td class="p-2 border font-semibold bg-slate-50">${pos}</td>
                        <td class="p-1 border"><input type="text" class="w-full border rounded p-1 step1-input" placeholder="Brand"></td>
                        <td class="p-1 border"><input type="text" class="w-full border rounded p-1 step1-input" placeholder="Size"></td>
                        <td class="p-1 border"><input type="number" step="0.1" class="w-full border rounded p-1 step1-input" placeholder="mm"></td>
                    </tr>`;
                });
                container.innerHTML = html + `</table>`;
            }

            function autoCalculateFinishDate() {
                const input = document.getElementById('inspectionDate').value;
                if (!input) return;
                let date = input.includes("-") ? new Date(input) : new Date(Date.parse(input));
                if (isNaN(date.getTime())) return;
                date.setDate(date.getDate() + 42); 
                if (date.getDay() === 6) date.setDate(date.getDate() + 2);
                if (date.getDay() === 0) date.setDate(date.getDate() + 1);
                const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                document.getElementById('finishDate').value = `${String(date.getDate()).padStart(2, '0')} ${months[date.getMonth()]} ${date.getFullYear()}`;
            }

            function enableEditStep1() {
                document.querySelectorAll('.step1-input').forEach(i => i.removeAttribute('disabled'));
                document.getElementById('topSaveBtn').classList.remove('hidden');
                document.getElementById('mainSaveBtn').classList.remove('hidden');
                document.getElementById('topEditBtn').classList.add('hidden');
            }

            function goToStep2() {
                document.getElementById('step1Page').classList.add('hidden');
                document.getElementById('step2Page').classList.remove('hidden');
                document.getElementById('backBtn').classList.remove('hidden');
                document.getElementById('topEditBtn').classList.add('hidden');
                document.getElementById('topSaveBtn').classList.add('hidden');
                if (isStep2Locked) document.getElementById('topEditStep2Btn').classList.remove('hidden');
                else document.getElementById('topSaveStep2Btn').classList.remove('hidden');
            }

            function goBackToStep1() {
                document.getElementById('step2Page').classList.add('hidden');
                document.getElementById('step1Page').classList.remove('hidden');
                document.getElementById('backBtn').classList.add('hidden');
                document.getElementById('topSaveStep2Btn').classList.add('hidden');
                document.getElementById('topEditStep2Btn').classList.add('hidden');
                document.getElementById('topEditBtn').classList.remove('hidden');
            }

            function confirmExit() { if (confirm("确定要退出至初启画面吗？")) alert("回到初启主页"); }

            function unlockStep2() {
                isStep2Locked = false;
                document.querySelectorAll('#step2Page input, #step2Page select, #step2Page textarea').forEach(el => {
                    if (!el.readOnly) el.removeAttribute('disabled');
                });
                document.querySelectorAll('.step2-btn, .del-btn').forEach(btn => btn.classList.remove('hidden'));
                document.getElementById('topSaveStep2Btn').classList.remove('hidden');
                document.getElementById('topEditStep2Btn').classList.add('hidden');
                document.getElementById('bottomSaveStep2Btn').classList.remove('hidden');
                document.getElementById('bottomEditStep2Btn').classList.add('hidden');
            }

            let partsRowCounter = 0;
            function addPartsRow() {
                partsRowCounter++;
                const tbody = document.getElementById('partsTbody');
                const tr = document.createElement('tr');
                tr.className = "border-b hover:bg-slate-50 parts-row";
                tr.id = `partRow_${partsRowCounter}`;
                tr.innerHTML = `
                    <td class="p-2 border text-center font-bold text-gray-500 row-sn"></td>
                    <td class="p-1 border"><input type="number" value="1" min="1" onchange="calcPartsRow(${partsRowCounter})" onkeyup="calcPartsRow(${partsRowCounter})" id="partQty_${partsRowCounter}" class="w-full border rounded p-1 text-xs text-center font-semibold step2-input"></td>
                    <td class="p-1 border">
                        <input type="text" id="partNameInput_${partsRowCounter}" placeholder="Type part & press Enter..." onblur="fetchPartPriceByInput(${partsRowCounter})" onkeydown="if(event.key==='Enter'){this.blur();}" class="w-full border rounded p-1 text-xs bg-white step2-input font-medium">
                    </td>
                    <td class="p-1 border"><input type="text" placeholder="e.g. de / c2" onkeyup="calcPartsRow(${partsRowCounter})" id="partCode_${partsRowCounter}" class="w-full border rounded p-1 text-xs font-mono uppercase font-bold text-blue-600 step2-input"></td>
                    <td class="p-1 border bg-slate-50"><input type="text" readonly id="partCondition_${partsRowCounter}" class="w-full p-1 text-xs bg-transparent font-medium text-slate-700 outline-none cursor-default" placeholder="-"></td>
                    <td class="p-1 border"><input type="text" placeholder="0.00" oninput="formatNumberInput(this)" onchange="calcPartsRow(${partsRowCounter})" onkeyup="calcPartsRow(${partsRowCounter})" id="partNett_${partsRowCounter}" class="w-full border rounded p-1 text-xs font-mono text-right step2-input font-bold text-slate-700"></td>
                    <td class="p-1 border bg-slate-50"><input type="text" readonly id="partEst_${partsRowCounter}" class="w-full p-1 text-xs font-mono text-right bg-transparent text-slate-800 outline-none font-bold" placeholder="0.00"></td>
                    <td class="p-1 border bg-slate-50"><input type="text" readonly id="partAss_${partsRowCounter}" class="w-full p-1 text-xs font-mono text-right font-bold bg-transparent text-slate-800 outline-none" placeholder="0.00"></td>
                    <td class="p-1 border text-center"><button onclick="removePartsRow(${partsRowCounter})" class="del-btn text-red-500 hover:text-red-700 font-bold px-1">✕</button></td>
                `;
                tbody.appendChild(tr);
                updatePartsSN();
            }

            function removePartsRow(rowId) {
                const row = document.getElementById(`partRow_${rowId}`);
                if (row) { row.remove(); updatePartsSN(); calcPartsTotals(); }
            }

            function updatePartsSN() {
                document.querySelectorAll('#partsTbody .parts-row').forEach((row, idx) => { row.querySelector('.row-sn').innerText = idx + 1; });
            }

            function recalcAllPartsRows() {
                document.querySelectorAll('#partsTbody .parts-row').forEach(row => { calcPartsRow(row.id.split('_')[1]); });
            }

            function calcPartsRow(rowId) {
                const qty = parseNumber(document.getElementById(`partQty_${rowId}`).value);
                const nettInput = document.getElementById(`partNett_${rowId}`);
                const nettPrice = parseNumber(nettInput.value);
                const partName = document.getElementById(`partNameInput_${rowId}`).value.trim();
                const code = document.getElementById(`partCode_${rowId}`).value.trim().toLowerCase();

                if (partName !== "" && (nettInput.value === "" || nettPrice === 0)) {
                    nettInput.classList.add("bg-red-200", "text-red-900", "border-red-400");
                } else {
                    nettInput.classList.remove("bg-red-200", "text-red-900", "border-red-400");
                }

                const globalPctInput = parseNumber(document.getElementById('globalPercentage').value);
                const estimateTotal = nettPrice * qty * (globalPctInput / 100.0);
                document.getElementById(`partEst_${rowId}`).value = formatNumber(estimateTotal);

                let conditionWording = "-";
                let conditionAssMultiplier = 1.0; 

                if (code in remarkRules) {
                    const rule = remarkRules[code];
                    conditionWording = rule.wording;
                    if (rule.type === "zero") conditionAssMultiplier = 0;
                    else if (rule.type === "full") conditionAssMultiplier = 1.0;
                    else if (rule.type === "m_pcs") conditionAssMultiplier = qty > 0 ? (1.0 / qty) * rule.pcs : 0;
                    else if (rule.type === "cut") conditionAssMultiplier = 1.0 - rule.percent;
                } else if (code === "") {
                    conditionWording = "-"; 
                    conditionAssMultiplier = 1.0;
                } else {
                    conditionWording = "UNKNOWN CODE"; 
                    conditionAssMultiplier = 0;
                }

                document.getElementById(`partCondition_${rowId}`).value = conditionWording;
                document.getElementById(`partAss_${rowId}`).value = formatNumber(estimateTotal * conditionAssMultiplier);
                calcPartsTotals();
            }

            function calcPartsTotals() {
                let subEst = 0, subAss = 0;
                document.querySelectorAll('#partsTbody .parts-row').forEach(row => {
                    const rId = row.id.split('_')[1];
                    subEst += parseNumber(document.getElementById(`partEst_${rId}`).value);
                    subAss += parseNumber(document.getElementById(`partAss_${rId}`).value);
                });
                const discEst = subEst * 0.10, discAss = subAss * 0.10;
                const totalEst = subEst - discEst, totalAss = subAss - discAss;

                document.getElementById('partsSubEst').innerText = formatNumber(subEst);
                document.getElementById('partsSubAss').innerText = formatNumber(subAss);
                document.getElementById('partsDiscEst').innerText = formatNumber(discEst);
                document.getElementById('partsDiscAss').innerText = formatNumber(discAss);
                document.getElementById('partsTotalEst').innerText = formatNumber(totalEst);
                document.getElementById('partsTotalAss').innerText = formatNumber(totalAss);
                calcGrandTotals();
            }

            let miscRowCounter = 0;
            function addMiscRow() {
                miscRowCounter++;
                const tbody = document.getElementById('miscTbody');
                const tr = document.createElement('tr');
                tr.className = "border-b hover:bg-slate-50 misc-row";
                tr.id = `miscRow_${miscRowCounter}`;
                tr.innerHTML = `
                    <td class="p-2 border text-center font-bold text-gray-500 misc-sn"></td>
                    <td class="p-1 border"><input type="number" value="1" min="1" id="miscQty_${miscRowCounter}" class="w-full border rounded p-1 text-xs text-center font-semibold step2-input"></td>
                    <td class="p-1 border"><input type="text" placeholder="Misc description..." class="w-full border rounded p-1 text-xs step2-input"></td>
                    <td class="p-1 border"><input type="text" placeholder="Code..." onkeyup="updateMiscCondition(${miscRowCounter})" id="miscCode_${miscRowCounter}" class="w-full border rounded p-1 text-xs font-mono uppercase font-bold text-blue-600 step2-input"></td>
                    <td class="p-1 border bg-slate-50"><input type="text" readonly id="miscCondition_${miscRowCounter}" class="w-full p-1 text-xs bg-transparent font-medium text-slate-700 outline-none" placeholder="-"></td>
                    <td class="p-1 border"><input type="text" placeholder="0.00" oninput="formatNumberInput(this)" onkeyup="calcMiscTotals()" id="miscEst_${miscRowCounter}" class="w-full border rounded p-1 text-xs font-mono text-right step2-input"></td>
                    <td class="p-1 border"><input type="text" placeholder="0.00" oninput="formatNumberInput(this)" onkeyup="calcMiscTotals()" id="miscAss_${miscRowCounter}" class="w-full border rounded p-1 text-xs font-mono text-right font-bold text-slate-800 step2-input"></td>
                    <td class="p-1 border text-center"><button onclick="removeMiscRow(${miscRowCounter})" class="del-btn text-red-500 hover:text-red-700 font-bold px-1">✕</button></td>
                `;
                tbody.appendChild(tr);
                updateMiscSN();
            }

            function updateMiscCondition(rowId) {
                const code = document.getElementById(`miscCode_${rowId}`).value.trim().toLowerCase();
                let wording = "-";
                if (code in remarkRules) wording = remarkRules[code].wording;
                else if (code !== "") wording = "UNKNOWN CODE";
                document.getElementById(`miscCondition_${rowId}`).value = wording;
            }

            function removeMiscRow(rowId) {
                const row = document.getElementById(`miscRow_${rowId}`);
                if (row) { row.remove(); updateMiscSN(); calcMiscTotals(); }
            }

            function updateMiscSN() {
                document.querySelectorAll('#miscTbody .misc-row').forEach((row, idx) => { row.querySelector('.misc-sn').innerText = idx + 1; });
            }

            function calcMiscTotals() {
                let totalEst = 0, totalAss = 0;
                document.querySelectorAll('#miscTbody .misc-row').forEach(row => {
                    const rId = row.id.split('_')[1];
                    totalEst += parseNumber(document.getElementById(`miscEst_${rId}`).value);
                    totalAss += parseNumber(document.getElementById(`miscAss_${rId}`).value);
                });
                document.getElementById('miscTotalEst').innerText = formatNumber(totalEst);
                document.getElementById('miscTotalAss').innerText = formatNumber(totalAss);
                calcGrandTotals();
            }

            const labourItems = [
                "To remove, cut out damage portion, jack out, straighten, panel beating, welding, align and renew replaced parts.",
                "To check wirings and lightings.",
                "To remove & install reverse sensor.",
                "To remove & refix tailgate fittings to facilitate repair.",
                "To remove & refix rear upholstery, garnish and attachment.",
                "To supply sealant and to seal off all weld spot seam and gaps.",
                "To spray anti-rust coating.",
                "To remove and install rear windscreen glass.",
                "To remove and install front windscreen glass.",
                "To remove and install rear right quarter glass.",
                "To remove and install rear left quarter glass.",
                "To remove and install front right quarter glass.",
                "To remove and install front left quarter glass.",
                "To recalibrate right power gate sliding door and use diagnostic computer to reset the system.",
                "To recalibrate left power gate sliding door and use diagnostic computer to reset the system.",
                "To change radiator and related parts , and to perform coolant leakage test .",
                "To change a/c parts and to recharge a/c gas.",
                "To conduct wheel alignment.",
                "To transfer door glass, trimboard and other components to replacement door.",
                "To remove and install rear suspension unit.",
                "To remove and install front suspension unit.",
                "To remove and install front left suspension unit.",
                "To remove and install front right suspension unit.",
                "To remove and install rear left suspension unit.",
                "To remove and install rear right suspension unit.",
                "To remove and install exhaust pipes."
            ];

            function initLabourChecklist() {
                const container = document.getElementById('labourChecklist');
                labourItems.forEach((item, idx) => {
                    const div = document.createElement('div');
                    div.id = `labourRow_${idx}`;
                    div.className = "flex items-center justify-between p-2 rounded bg-slate-100 border text-xs text-slate-500";
                    div.innerHTML = `
                        <div class="flex items-center space-x-2 flex-1 mr-4">
                            <input type="checkbox" onchange="toggleLabourRow(${idx}, this.checked)" class="step2-input w-4 h-4 rounded text-blue-600 focus:ring-0">
                            <span class="font-medium">${item}</span>
                        </div>
                        <div class="flex space-x-2 hidden" id="labourInputs_${idx}">
                            <input type="text" placeholder="Est (S$)" oninput="formatNumberInput(this)" onkeyup="calcLabourTotals()" id="labourEst_${idx}" class="step2-input w-24 border rounded p-1 bg-white text-slate-800 font-mono">
                            <input type="text" placeholder="Ass (S$)" oninput="formatNumberInput(this)" onkeyup="calcLabourTotals()" id="labourAss_${idx}" class="step2-input w-24 border rounded p-1 bg-white text-slate-800 font-mono font-bold">
                        </div>
                    `;
                    container.appendChild(div);
                });
            }

            function toggleLabourRow(idx, checked) {
                const row = document.getElementById(`labourRow_${idx}`);
                const inputs = document.getElementById(`labourInputs_${idx}`);
                if (checked) {
                    row.className = "flex items-center justify-between p-2 rounded bg-blue-50 border-blue-300 text-xs text-blue-900 font-bold shadow-sm";
                    inputs.classList.remove('hidden');
                } else {
                    row.className = "flex items-center justify-between p-2 rounded bg-slate-100 border text-xs text-slate-500";
                    inputs.classList.add('hidden');
                    document.getElementById(`labourEst_${idx}`).value = "";
                    document.getElementById(`labourAss_${idx}`).value = "";
                }
                calcLabourTotals();
            }

            let customLabourCounter = 0;
            function addCustomLabourRow() {
                customLabourCounter++;
                const container = document.getElementById('customLabourContainer');
                const div = document.createElement('div');
                div.className = "flex items-center justify-between p-2 rounded bg-white border text-xs shadow-sm custom-labour-row";
                div.id = `customLabourRow_${customLabourCounter}`;
                div.innerHTML = `
                    <div class="flex-1 mr-4"><input type="text" placeholder="Custom labour description..." class="step2-input w-full border rounded p-1 text-xs"></div>
                    <div class="flex space-x-2 items-center">
                        <input type="text" placeholder="Est (S$)" oninput="formatNumberInput(this)" onkeyup="calcLabourTotals()" id="customLabourEst_${customLabourCounter}" class="step2-input w-24 border rounded p-1 bg-white font-mono">
                        <input type="text" placeholder="Ass (S$)" oninput="formatNumberInput(this)" onkeyup="calcLabourTotals()" id="customLabourAss_${customLabourCounter}" class="step2-input w-24 border rounded p-1 bg-white font-mono font-bold">
                        <button onclick="removeCustomLabourRow(${customLabourCounter})" class="del-btn text-red-500 hover:text-red-700 font-bold px-1 text-sm">✕</button>
                    </div>
                `;
                container.appendChild(div);
            }

            function removeCustomLabourRow(id) {
                const row = document.getElementById(`customLabourRow_${id}`);
                if(row) { row.remove(); calcLabourTotals(); }
            }

            function calcLabourTotals() {
                let totalEst = 0, totalAss = 0;
                labourItems.forEach((_, idx) => {
                    totalEst += parseNumber(document.getElementById(`labourEst_${idx}`)?.value);
                    totalAss += parseNumber(document.getElementById(`labourAss_${idx}`)?.value);
                });
                document.querySelectorAll('.custom-labour-row').forEach(row => {
                    const cId = row.id.split('_')[1];
                    totalEst += parseNumber(document.getElementById(`customLabourEst_${cId}`)?.value);
                    totalAss += parseNumber(document.getElementById(`customLabourAss_${cId}`)?.value);
                });
                document.getElementById('labourTotalEst').innerText = formatNumber(totalEst);
                document.getElementById('labourTotalAss').innerText = formatNumber(totalAss);
                calcGrandTotals();
            }

            function calcGrandTotals() {
                const partsEst = parseNumber(document.getElementById('partsTotalEst').innerText);
                const partsAss = parseNumber(document.getElementById('partsTotalAss').innerText);
                const miscEst = parseNumber(document.getElementById('miscTotalEst').innerText);
                const miscAss = parseNumber(document.getElementById('miscTotalAss').innerText);
                const labourEst = parseNumber(document.getElementById('labourTotalEst').innerText);
                const labourAss = parseNumber(document.getElementById('labourTotalAss').innerText);

                const paintLabourEst = parseNumber(document.getElementById('paintLabourEst').value);
                const paintLabourAss = parseNumber(document.getElementById('paintLabourAss').value);
                const paintMatEst = parseNumber(document.getElementById('paintMatEst').value);
                const paintMatAss = parseNumber(document.getElementById('paintMatAss').value);

                document.getElementById('finalPartsEst').innerText = "S$ " + formatNumber(partsEst);
                document.getElementById('finalPartsAss').innerText = "S$ " + formatNumber(partsAss);
                document.getElementById('finalMiscEst').innerText = "S$ " + formatNumber(miscEst);
                document.getElementById('finalMiscAss').innerText = "S$ " + formatNumber(miscAss);
                document.getElementById('finalLabourEst').innerText = "S$ " + formatNumber(labourEst);
                document.getElementById('finalLabourAss').innerText = "S$ " + formatNumber(labourAss);
                document.getElementById('finalPaintLabourEst').innerText = "S$ " + formatNumber(paintLabourEst);
                document.getElementById('finalPaintLabourAss').innerText = "S$ " + formatNumber(paintLabourAss);
                document.getElementById('finalPaintMatEst').innerText = "S$ " + formatNumber(paintMatEst);
                document.getElementById('finalPaintMatAss').innerText = "S$ " + formatNumber(paintMatAss);

                document.getElementById('finalTotalEst').innerText = "S$ " + formatNumber(partsEst + miscEst + labourEst + paintLabourEst + paintMatEst);
                document.getElementById('finalTotalAss').innerText = "S$ " + formatNumber(partsAss + miscAss + labourAss + paintLabourAss + paintMatAss);
            }

            window.onload = function() {
                addPartsRow();
                addMiscRow();
                initLabourChecklist();
            };
        </script>
    </body>
    </html>
    """
    return html_content

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)