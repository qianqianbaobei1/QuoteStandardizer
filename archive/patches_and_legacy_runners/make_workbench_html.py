# -*- coding: utf-8 -*-
"""
生成《底稿超清查验工作台.html》
单文件纯离线 HTML 工具，左右分栏、无限缩放、拖拽平移、图表联动
"""
import os, glob, json, openpyxl

vendors_data = {}
excel_files = sorted(glob.glob("deliveries/*-标准解析.xlsx"))

for ef in excel_files[:10]: # Include top 10 representative vendors with sample rows
    vname = os.path.basename(ef).replace("-标准解析.xlsx", "")
    img_dir = f"deliveries/{vname}_图片库"
    if not os.path.exists(img_dir):
        continue
    imgs = sorted([f for f in os.listdir(img_dir) if f.endswith(".png") and not f.startswith("thumb_")])
    
    wb = openpyxl.load_workbook(ef, read_only=True)
    ws = wb["标准化报价"]
    headers = [c for c in next(ws.iter_rows(values_only=True))]
    
    # Extract sample rows (up to 30 rows per vendor)
    rows_sample = []
    idx_name = headers.index("材料名称") if "材料名称" in headers else 3
    idx_spec = headers.index("规格参数") if "规格参数" in headers else 4
    idx_unit = headers.index("单位") if "单位" in headers else 5
    idx_price = headers.index("含税出厂价") if "含税出厂价" in headers else 6
    idx_page = headers.index("厂家报价文件页码") if "厂家报价文件页码" in headers else 21
    idx_img = headers.index("厂家报价图片文件名") if "厂家报价图片文件名" in headers else 22
    
    count = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        count += 1
        if count <= 25 or (count % 1000 == 0 and len(rows_sample) < 50):
            rows_sample.append({
                "row_num": count,
                "name": str(r[idx_name] or ""),
                "spec": str(r[idx_spec] or "").replace("\n", " "),
                "unit": str(r[idx_unit] or ""),
                "price": str(r[idx_price] or ""),
                "page": str(r[idx_page] or ""),
                "img": str(r[idx_img] or (imgs[0] if imgs else ""))
            })
    wb.close()
    
    vendors_data[vname] = {
        "images": [f"../{vname}_图片库/{img}" for img in imgs],
        "image_names": imgs,
        "rows": rows_sample
    }

html_content = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>建材报价单底稿超清查验工作台</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Microsoft YaHei", sans-serif; background: #0f172a; color: #e2e8f0; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
        header { background: #1e293b; padding: 12px 24px; display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #334155; }
        .logo-area { display: flex; align-items: center; gap: 12px; }
        .badge { background: #0284c7; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; letter-spacing: 0.5px; }
        h1 { font-size: 17px; font-weight: 600; color: #f8fafc; }
        .controls { display: flex; align-items: center; gap: 14px; }
        select { background: #0f172a; color: #f8fafc; border: 1px solid #475569; padding: 6px 12px; border-radius: 6px; font-size: 13px; outline: none; }
        .btn { background: #334155; color: white; border: none; padding: 6px 14px; border-radius: 6px; font-size: 13px; cursor: pointer; transition: 0.2s; display: flex; align-items: center; gap: 6px; }
        .btn:hover { background: #475569; }
        .btn-primary { background: #0284c7; }
        .btn-primary:hover { background: #0369a1; }
        
        main { display: flex; flex: 1; overflow: hidden; }
        .panel-left { width: 48%; background: #1e293b; border-right: 1px solid #334155; display: flex; flex-direction: column; }
        .panel-right { width: 52%; background: #090d16; display: flex; flex-direction: column; position: relative; }
        
        .panel-title { padding: 10px 16px; background: #182234; border-bottom: 1px solid #334155; font-size: 13px; font-weight: 600; color: #94a3b8; display: flex; justify-content: space-between; align-items: center; }
        .table-container { flex: 1; overflow-y: auto; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; }
        th { background: #182234; color: #94a3b8; padding: 8px 12px; font-weight: 600; position: sticky; top: 0; border-bottom: 1px solid #334155; white-space: nowrap; }
        td { padding: 8px 12px; border-bottom: 1px solid #243248; color: #cbd5e1; white-space: nowrap; max-width: 240px; overflow: hidden; text-overflow: ellipsis; }
        tr:hover { background: #26354a; cursor: pointer; }
        tr.active { background: #0369a1 !important; color: white; }
        .img-tag { color: #38bdf8; text-decoration: underline; }
        
        .image-toolbar { padding: 8px 16px; background: #182234; border-bottom: 1px solid #334155; display: flex; justify-content: space-between; align-items: center; }
        .zoom-info { font-size: 12px; color: #94a3b8; }
        .viewer-container { flex: 1; overflow: hidden; position: relative; cursor: grab; background: radial-gradient(#1e293b 10%, transparent 10%); background-size: 20px 20px; }
        .viewer-container:active { cursor: grabbing; }
        .master-image { position: absolute; transform-origin: 0 0; transition: transform 0.05s ease-out; box-shadow: 0 10px 30px rgba(0,0,0,0.8); }
        .page-tabs { display: flex; gap: 6px; overflow-x: auto; max-width: 400px; }
        .page-tab { background: #0f172a; border: 1px solid #334155; color: #94a3b8; padding: 3px 10px; border-radius: 4px; font-size: 11px; cursor: pointer; white-space: nowrap; }
        .page-tab.active { background: #0284c7; color: white; border-color: #0284c7; font-weight: bold; }
    </style>
</head>
<body>
    <header>
        <div class="logo-area">
            <span class="badge">PRO 4K</span>
            <h1>建材报价单底稿超清查验工作台</h1>
        </div>
        <div class="controls">
            <label for="vendorSelect" style="font-size: 13px; color: #94a3b8;">选择供应商：</label>
            <select id="vendorSelect"></select>
            <button class="btn btn-primary" onclick="resetZoom()">1:1 原始尺寸</button>
            <button class="btn" onclick="fitScreen()">适应屏幕</button>
        </div>
    </header>
    <main>
        <div class="panel-left">
            <div class="panel-title">
                <span>标准化解析报价表 (点击行直达底稿对应页)</span>
                <span id="rowStats" style="font-weight: normal; font-size: 12px; color: #64748b;"></span>
            </div>
            <div class="table-container">
                <table id="dataTable">
                    <thead>
                        <tr>
                            <th>序号</th>
                            <th>材料名称</th>
                            <th>规格参数 (已过滤编码)</th>
                            <th>单位</th>
                            <th>单价(元)</th>
                            <th>底稿页码</th>
                            <th>底稿图片</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>
        </div>
        <div class="panel-right">
            <div class="image-toolbar">
                <div class="page-tabs" id="pageTabs"></div>
                <div class="zoom-info">
                    <span id="zoomLevel">缩放: 100%</span> | 
                    <span>滚轮放大 / 拖拽平移</span>
                </div>
            </div>
            <div class="viewer-container" id="viewerContainer">
                <img id="masterImage" class="master-image" src="" alt="底稿加载中...">
            </div>
        </div>
    </main>
    <script>
        const data = """ + json.dumps(vendors_data, ensure_ascii=False) + """;
        const vendorSelect = document.getElementById('vendorSelect');
        const tableBody = document.getElementById('tableBody');
        const pageTabs = document.getElementById('pageTabs');
        const masterImage = document.getElementById('masterImage');
        const viewerContainer = document.getElementById('viewerContainer');
        const zoomLevelEl = document.getElementById('zoomLevel');
        const rowStats = document.getElementById('rowStats');

        let currentVendor = '';
        let currentImgIdx = 0;
        let scale = 1.0;
        let panX = 40;
        let panY = 40;
        let isDragging = false;
        let startX, startY;

        // Init select
        for (let v in data) {
            let opt = document.createElement('option');
            opt.value = v;
            opt.textContent = v;
            vendorSelect.appendChild(opt);
        }
        currentVendor = Object.keys(data)[0];
        vendorSelect.addEventListener('change', (e) => {
            currentVendor = e.target.value;
            loadVendor(currentVendor);
        });

        function loadVendor(vendor) {
            let vData = data[vendor];
            currentImgIdx = 0;
            rowStats.textContent = `抽样展示 ${vData.rows.length} 行`;
            
            // Build tabs
            pageTabs.innerHTML = '';
            vData.image_names.forEach((name, idx) => {
                let tab = document.createElement('div');
                tab.className = 'page-tab' + (idx === 0 ? ' active' : '');
                tab.textContent = name.replace('quote_img_', '').replace('quote_img', '页码 ');
                tab.onclick = () => switchPage(idx);
                pageTabs.appendChild(tab);
            });

            // Build table
            tableBody.innerHTML = '';
            vData.rows.forEach((r, idx) => {
                let tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${r.row_num}</td>
                    <td><strong>${r.name}</strong></td>
                    <td title="${r.spec}">${r.spec.substring(0, 32)}</td>
                    <td>${r.unit}</td>
                    <td style="color: #38bdf8; font-weight: bold;">${r.price}</td>
                    <td>${r.page}</td>
                    <td class="img-tag">${r.img}</td>
                `;
                tr.onclick = () => {
                    document.querySelectorAll('tr').forEach(el => el.classList.remove('active'));
                    tr.classList.add('active');
                    let targetImg = r.img;
                    let targetIdx = vData.image_names.indexOf(targetImg);
                    if (targetIdx >= 0) switchPage(targetIdx);
                };
                tableBody.appendChild(tr);
            });

            switchPage(0);
        }

        function switchPage(idx) {
            let vData = data[currentVendor];
            if (idx >= vData.images.length) idx = 0;
            currentImgIdx = idx;
            document.querySelectorAll('.page-tab').forEach((t, i) => {
                t.className = 'page-tab' + (i === idx ? ' active' : '');
            });
            masterImage.src = vData.images[idx];
            masterImage.onload = () => {
                fitScreen();
            };
        }

        function updateTransform() {
            masterImage.style.transform = `translate(${panX}px, ${panY}px) scale(${scale})`;
            zoomLevelEl.textContent = `缩放: ${Math.round(scale * 100)}%`;
        }

        function fitScreen() {
            let cw = viewerContainer.clientWidth;
            let ch = viewerContainer.clientHeight;
            let iw = masterImage.naturalWidth || 1000;
            let ih = masterImage.naturalHeight || 1000;
            scale = Math.min((cw - 60) / iw, (ch - 60) / ih);
            panX = (cw - iw * scale) / 2;
            panY = 20;
            updateTransform();
        }

        function resetZoom() {
            scale = 1.0;
            panX = 40;
            panY = 40;
            updateTransform();
        }

        // Pan & Zoom
        viewerContainer.addEventListener('wheel', (e) => {
            e.preventDefault();
            let zoomFactor = e.deltaY < 0 ? 1.15 : 0.85;
            let newScale = Math.min(Math.max(scale * zoomFactor, 0.1), 10.0);
            
            let rect = viewerContainer.getBoundingClientRect();
            let mouseX = e.clientX - rect.left;
            let mouseY = e.clientY - rect.top;

            panX = mouseX - (mouseX - panX) * (newScale / scale);
            panY = mouseY - (mouseY - panY) * (newScale / scale);
            scale = newScale;
            updateTransform();
        }, { passive: false });

        viewerContainer.addEventListener('mousedown', (e) => {
            isDragging = true;
            startX = e.clientX - panX;
            startY = e.clientY - panY;
        });

        window.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            panX = e.clientX - startX;
            panY = e.clientY - startY;
            updateTransform();
        });

        window.addEventListener('mouseup', () => { isDragging = false; });

        // Start
        loadVendor(currentVendor);
    </script>
</body>
</html>
"""

out_html = "deliveries/技术方案/底稿超清查验工作台.html"
with open(out_html, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Successfully generated standalone Web workbench: {out_html}")
