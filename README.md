<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>INCU Analyzer</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mqtt/4.3.7/mqtt.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root {
            --bg:        #f5f3ef;
            --bg2:       #ffffff;
            --bg3:       #eeeae4;
            --surface:   #faf9f7;
            --border:    #e8e3db;
            --border2:   #ddd8cf;

            --teal:      #4a9b8e;
            --teal-lt:   #e8f5f3;
            --teal-mid:  #b8ddd9;
            --sage:      #7a9e7e;
            --sage-lt:   #edf4ee;
            --blush:     #e07a7a;
            --blush-lt:  #fceaea;
            --sand:      #c4956a;
            --sand-lt:   #fdf3ea;
            --lavender:  #8b7ec8;
            --lav-lt:    #f0eefb;
            --sky:       #5b9bd5;
            --sky-lt:    #eaf3fb;

            --text:      #3d3530;
            --text-2:    #7a726b;
            --text-3:    #a89f97;

            --radius:    16px;
            --radius-sm: 10px;
            --shadow:    0 2px 12px rgba(60,50,40,0.07);
            --shadow-md: 0 6px 28px rgba(60,50,40,0.10);

            --mono: 'DM Mono', monospace;
            --sans: 'DM Sans', sans-serif;
        }

        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: var(--sans);
            background: var(--bg);
            color: var(--text);
            min-height: 100vh;
        }

        /* ── SIDEBAR STRIP ── */
        .sidebar-strip {
            position: fixed;
            left: 0; top: 0; bottom: 0;
            width: 6px;
            background: linear-gradient(180deg, var(--teal) 0%, var(--sage) 50%, var(--lavender) 100%);
            z-index: 100;
        }

        /* ── PAGE ── */
        .page {
            margin-left: 6px;
            padding: 36px 40px 60px;
            max-width: 1380px;
        }

        /* ── HEADER ── */
        header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 36px;
            flex-wrap: wrap;
            gap: 16px;
        }

        .brand {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .brand-icon {
            width: 50px; height: 50px;
            background: linear-gradient(135deg, var(--teal), var(--sage));
            border-radius: 14px;
            display: flex; align-items: center; justify-content: center;
            font-size: 1.4em;
            box-shadow: 0 4px 14px rgba(74,155,142,0.3);
        }

        .brand-text h1 {
            font-size: 1.55em;
            font-weight: 600;
            color: var(--text);
            letter-spacing: -0.02em;
        }

        .brand-text p {
            font-size: 0.78em;
            color: var(--text-3);
            font-weight: 400;
            margin-top: 1px;
        }

        .mqtt-badge {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 9px 16px;
            border-radius: 100px;
            background: var(--bg2);
            border: 1.5px solid var(--border);
            font-size: 0.8em;
            font-weight: 500;
            color: var(--text-2);
            box-shadow: var(--shadow);
            transition: all 0.4s;
        }

        .mqtt-badge .dot {
            width: 8px; height: 8px;
            border-radius: 50%;
            background: var(--blush);
            transition: all 0.4s;
        }

        .mqtt-badge.connected {
            border-color: var(--teal-mid);
            color: var(--teal);
            background: var(--teal-lt);
        }

        .mqtt-badge.connected .dot {
            background: var(--teal);
            animation: pulse-dot 2s infinite;
        }

        @keyframes pulse-dot {
            0%, 100% { box-shadow: 0 0 0 0 rgba(74,155,142,0.4); }
            50%       { box-shadow: 0 0 0 5px rgba(74,155,142,0); }
        }

        /* ── CONTROLS CARD ── */
        .controls-card {
            background: var(--bg2);
            border: 1.5px solid var(--border);
            border-radius: var(--radius);
            padding: 24px 28px;
            box-shadow: var(--shadow);
            margin-bottom: 28px;
            display: flex;
            align-items: flex-end;
            gap: 20px;
            flex-wrap: wrap;
        }

        .field {
            display: flex;
            flex-direction: column;
            gap: 6px;
        }

        .field label {
            font-size: 0.73em;
            font-weight: 500;
            color: var(--text-3);
            letter-spacing: 0.04em;
            text-transform: uppercase;
        }

        .field input {
            font-family: var(--mono);
            font-size: 0.95em;
            padding: 10px 14px;
            background: var(--surface);
            border: 1.5px solid var(--border);
            border-radius: var(--radius-sm);
            color: var(--text);
            width: 150px;
            outline: none;
            transition: border-color 0.25s, box-shadow 0.25s;
        }

        .field input:focus {
            border-color: var(--teal);
            box-shadow: 0 0 0 4px rgba(74,155,142,0.12);
        }

        .field input:disabled {
            opacity: 0.5;
            background: var(--bg3);
        }

        .btn {
            font-family: var(--sans);
            font-size: 0.88em;
            font-weight: 600;
            padding: 10px 22px;
            border-radius: 100px;
            border: 1.5px solid transparent;
            cursor: pointer;
            transition: all 0.22s;
            display: inline-flex;
            align-items: center;
            gap: 7px;
        }

        .btn-primary {
            background: var(--teal);
            color: #fff;
            border-color: var(--teal);
            box-shadow: 0 3px 12px rgba(74,155,142,0.3);
        }

        .btn-primary:hover {
            background: #3d8a7e;
            box-shadow: 0 5px 18px rgba(74,155,142,0.4);
            transform: translateY(-1px);
        }

        .btn-primary.recording {
            background: var(--blush);
            border-color: var(--blush);
            box-shadow: 0 3px 14px rgba(224,122,122,0.35);
            animation: rec-glow 1.5s infinite;
        }

        @keyframes rec-glow {
            0%, 100% { box-shadow: 0 3px 14px rgba(224,122,122,0.35); }
            50%       { box-shadow: 0 3px 22px rgba(224,122,122,0.6); }
        }

        .btn-ghost {
            background: transparent;
            color: var(--text-2);
            border-color: var(--border2);
        }

        .btn-ghost:hover {
            background: var(--bg3);
            color: var(--text);
            border-color: var(--border2);
        }

        .btn-accent {
            background: var(--lav-lt);
            color: var(--lavender);
            border-color: #d4cfee;
        }

        .btn-accent:hover {
            background: #e6e3f7;
            transform: translateY(-1px);
        }

        .timer-chip {
            margin-left: auto;
            text-align: right;
        }

        .timer-chip .t-label {
            font-size: 0.72em;
            color: var(--text-3);
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 3px;
        }

        .timer-chip .t-value {
            font-family: var(--mono);
            font-size: 2em;
            font-weight: 500;
            color: var(--teal);
            letter-spacing: 0.04em;
        }

        /* ── SENSOR CARDS ── */
        .section-label {
            font-size: 0.75em;
            font-weight: 600;
            color: var(--text-3);
            text-transform: uppercase;
            letter-spacing: 0.1em;
            margin-bottom: 14px;
            margin-top: 4px;
        }

        .sensor-grid {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 14px;
            margin-bottom: 14px;
        }

        .sensor-grid.g4 {
            grid-template-columns: repeat(4, 1fr);
            margin-bottom: 32px;
        }

        .sensor-card {
            background: var(--bg2);
            border: 1.5px solid var(--border);
            border-radius: var(--radius);
            padding: 20px 18px 18px;
            box-shadow: var(--shadow);
            transition: transform 0.2s, box-shadow 0.2s, border-color 0.3s;
            position: relative;
            overflow: hidden;
        }

        .sensor-card::after {
            content: '';
            position: absolute;
            bottom: 0; left: 0; right: 0;
            height: 3px;
            border-radius: 0 0 var(--radius) var(--radius);
            transition: opacity 0.3s;
        }

        .sensor-card:hover {
            transform: translateY(-2px);
            box-shadow: var(--shadow-md);
        }

        /* color themes per card */
        .card-teal::after  { background: linear-gradient(90deg, var(--teal), var(--teal-mid)); }
        .card-sage::after  { background: linear-gradient(90deg, var(--sage), #b8d4bb); }
        .card-sky::after   { background: linear-gradient(90deg, var(--sky), #a8ccec); }
        .card-sand::after  { background: linear-gradient(90deg, var(--sand), #ddb98a); }
        .card-lav::after   { background: linear-gradient(90deg, var(--lavender), #c0b9e8); }

        .sensor-card .card-label {
            font-size: 0.7em;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            color: var(--text-3);
            margin-bottom: 10px;
        }

        .sensor-card .card-val {
            font-family: var(--mono);
            font-size: 1.75em;
            font-weight: 500;
            color: var(--text);
            line-height: 1;
        }

        .sensor-card .card-val .unit {
            font-size: 0.5em;
            color: var(--text-3);
            margin-left: 2px;
        }

        .sensor-card .card-name {
            font-size: 0.72em;
            color: var(--text-3);
            margin-top: 5px;
        }

        /* updated flash */
        .sensor-card.flash {
            animation: card-flash 0.7s ease-out;
        }

        @keyframes card-flash {
            0%   { background: var(--teal-lt); border-color: var(--teal-mid); }
            100% { background: var(--bg2);     border-color: var(--border); }
        }

        /* error state */
        .sensor-card.error {
            border-color: #f0c0c0;
            background: var(--blush-lt);
        }

        .sensor-card.error .card-val { color: var(--blush); }
        .sensor-card.error::after { background: linear-gradient(90deg, var(--blush), #f0a0a0); }

        /* ── TABLE ── */
        .table-card {
            background: var(--bg2);
            border: 1.5px solid var(--border);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            overflow: hidden;
            margin-bottom: 28px;
        }

        .table-head {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 18px 24px 16px;
            border-bottom: 1.5px solid var(--border);
        }

        .table-head-title {
            font-weight: 600;
            font-size: 0.92em;
            color: var(--text);
        }

        .row-badge {
            background: var(--teal-lt);
            color: var(--teal);
            border-radius: 100px;
            padding: 3px 12px;
            font-size: 0.78em;
            font-weight: 600;
        }

        .table-scroll { overflow-x: auto; width: 100%; }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.83em;
            min-width: 100%;
            table-layout: fixed;
        }

        thead th {
            padding: 11px 16px;
            text-align: left;
            font-weight: 600;
            font-size: 0.78em;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            color: var(--text-3);
            background: var(--surface);
            border-bottom: 1.5px solid var(--border);
            white-space: nowrap;
            width: calc(100% / 11);
        }

        tbody td {
            padding: 10px 16px;
            border-bottom: 1px solid var(--border);
            color: var(--text-2);
            font-family: var(--mono);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        tbody tr:last-child td { border-bottom: none; }

        tbody tr { transition: background 0.15s; }
        tbody tr:hover { background: var(--surface); }

        tbody tr.oot { background: var(--blush-lt); }
        tbody tr.oot td { color: var(--blush); }
        tbody tr.oot:hover { background: #fce0e0; }

        .new-row { animation: row-slide 0.4s ease-out; }

        @keyframes row-slide {
            from { opacity: 0; transform: translateY(-6px); background: var(--teal-lt); }
            to   { opacity: 1; transform: translateY(0); }
        }

        /* ── SCROLLBAR ── */
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 4px; }

        /* ── RESPONSIVE ── */
        @media (max-width: 1100px) {
            .sensor-grid { grid-template-columns: repeat(3, 1fr); }
            .sensor-grid.g4 { grid-template-columns: repeat(2, 1fr); }
        }

        @media (max-width: 700px) {
            .page { padding: 24px 20px 48px; }
            .sensor-grid { grid-template-columns: repeat(2, 1fr); }
            .controls-card { flex-direction: column; align-items: flex-start; }
            .timer-chip { margin-left: 0; }
        }
    </style>
</head>
<body>

<div class="sidebar-strip"></div>

<div class="page">

    <!-- HEADER -->
    <header>
        <div class="brand">
            <div class="brand-icon">🌡️</div>
            <div class="brand-text">
                <h1>INCU Analyzer</h1>
                <p>Incubator Parameter Monitoring System</p>
            </div>
        </div>
        <div id="mqttBadge" class="mqtt-badge disconnected">
            <div class="dot"></div>
            <span id="mqttText">Disconnected</span>
        </div>
    </header>

    <!-- CONTROLS -->
    <div class="controls-card">
        <div class="field">
            <label>Interval (seconds)</label>
            <input type="number" id="intervalInput" min="1" value="2">
        </div>
        <div class="field">
            <label>Duration (HH:MM:SS)</label>
            <input type="text" id="timerInput" value="00:01:00" placeholder="00:00:00">
        </div>
        <button class="btn btn-primary" id="saveBtn">
            <span id="saveBtnIcon">▶</span>
            <span id="saveBtnText">Start Recording</span>
        </button>
        <button class="btn btn-ghost" id="resetBtn">↺ Reset</button>
        <button class="btn btn-accent" id="exportBtn">⬇ Export Excel</button>
        <div class="timer-chip">
            <div class="t-label">Remaining</div>
            <div class="t-value" id="timerDisplay">00:00:00</div>
        </div>
    </div>

    <!-- SENSOR READINGS ROW 1: T1-T5 -->
    <div class="section-label">Temperature Sensors</div>
    <div class="sensor-grid">
        <div class="sensor-card card-teal" id="t1Box">
            <div class="card-label">T1</div>
            <div class="card-val" id="t1Value">00.00<span class="unit">°C</span></div>
            <div class="card-name">Temperature 1</div>
        </div>
        <div class="sensor-card card-teal" id="t2Box">
            <div class="card-label">T2</div>
            <div class="card-val" id="t2Value">00.00<span class="unit">°C</span></div>
            <div class="card-name">Temperature 2</div>
        </div>
        <div class="sensor-card card-teal" id="t3Box">
            <div class="card-label">T3</div>
            <div class="card-val" id="t3Value">00.00<span class="unit">°C</span></div>
            <div class="card-name">Temperature 3</div>
        </div>
        <div class="sensor-card card-teal" id="t4Box">
            <div class="card-label">T4</div>
            <div class="card-val" id="t4Value">00.00<span class="unit">°C</span></div>
            <div class="card-name">Temperature 4</div>
        </div>
        <div class="sensor-card card-sage" id="t5Box">
            <div class="card-label">T5</div>
            <div class="card-val" id="t5Value">00.00<span class="unit">°C</span></div>
            <div class="card-name">Reference Temp</div>
        </div>
    </div>

    <!-- SENSOR READINGS ROW 2: TM, Flow, RH, Noise -->
    <div class="section-label">Environmental Sensors</div>
    <div class="sensor-grid g4">
        <div class="sensor-card card-sand" id="tmBox">
            <div class="card-label">TM</div>
            <div class="card-val" id="tmValue">00.00<span class="unit">°C</span></div>
            <div class="card-name">Mattress Temp</div>
        </div>
        <div class="sensor-card card-sky" id="flowBox">
            <div class="card-label">Flow</div>
            <div class="card-val" id="flowValue">00.00<span class="unit">m/s</span></div>
            <div class="card-name">Airflow</div>
        </div>
        <div class="sensor-card card-lav" id="rhBox">
            <div class="card-label">RH</div>
            <div class="card-val" id="rhValue">00.00<span class="unit">%</span></div>
            <div class="card-name">Humidity</div>
        </div>
        <div class="sensor-card card-sage" id="noiseBox">
            <div class="card-label">Noise</div>
            <div class="card-val" id="noiseValue">00.00<span class="unit">dB</span></div>
            <div class="card-name">Sound Level</div>
        </div>
    </div>

    <!-- DATA TABLE -->
    <div class="table-card">
        <div class="table-head">
            <div class="table-head-title">Recorded Data</div>
            <div class="row-badge" id="rowBadge">0 rows</div>
        </div>
        <div class="table-scroll">
            <table>
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Time</th>
                        <th>T1 (°C)</th>
                        <th>T2 (°C)</th>
                        <th>T3 (°C)</th>
                        <th>T4 (°C)</th>
                        <th>T5 (°C)</th>
                        <th>TM (°C)</th>
                        <th>Humidity (%)</th>
                        <th>Airflow (m/s)</th>
                        <th>Noise (dB)</th>
                    </tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>

</div>

<script>
    /* ── MQTT ─────────────────────────────────── */
    const client = mqtt.connect('wss://broker.hivemq.com:8884/mqtt', {
        clean: true,
        connectTimeout: 4000,
        reconnectPeriod: 1000,
        clientId: 'incu_soft_' + Math.random().toString(16).substr(2,8),
        keepalive: 60,
        protocolVersion: 4
    });

    client.on('connect', () => {
        setBadge('connected', '🟢 Connected');
        client.subscribe('incu/sensors');
    });
    client.on('error',   () => setBadge('disconnected', '⚠️ Error'));
    client.on('offline', () => { setBadge('disconnected', 'Disconnected'); resetSensors(); });

    client.on('message', (t, msg) => {
        try {
            const d = JSON.parse(msg.toString());
            lastDataTime = Date.now();
            currentData = d;
            renderSensors(d);
            flashAll();
        } catch(e){}
    });

    function setBadge(cls, txt) {
        const b = document.getElementById('mqttBadge');
        b.className = 'mqtt-badge ' + cls;
        document.getElementById('mqttText').textContent = txt;
    }

    /* ── STATE ────────────────────────────────── */
    let isRec = false, timerInt, dataInt, remaining = 0;
    let tableData = [], currentData = {}, lastDataTime = Date.now();

    setInterval(() => { if (Date.now() - lastDataTime > 10000) resetSensors(); }, 5000);

    /* ── SENSORS ──────────────────────────────── */
    function f(v) { return (v != null && v !== undefined) ? parseFloat(v).toFixed(2) : '00.00'; }

    function renderSensors(d) {
        setVal('t1Value', f(d.t1), '°C');
        setVal('t2Value', f(d.t2), '°C');
        setVal('t3Value', f(d.t3), '°C');
        setVal('t4Value', f(d.t4), '°C');
        setVal('t5Value', f(d.t5), '°C');
        setVal('tmValue', f(d.tm), '°C');
        setVal('flowValue', f(d.flow), 'm/s');
        setVal('rhValue', f(d.rh), '%');
        setVal('noiseValue', f(d.noise), 'dB');

        const t5 = parseFloat(d.t5 || 0);
        ['t1','t2','t3','t4'].forEach(k => {
            const v = parseFloat(d[k] || 0);
            const box = document.getElementById(k + 'Box');
            const isErr = v > 0 && t5 > 0 && (v < t5 - 0.8 || v > t5 + 0.8);
            box.classList.toggle('error', isErr);
        });
    }

    function setVal(id, val, unit) {
        document.getElementById(id).innerHTML = val + '<span class="unit"> ' + unit + '</span>';
    }

    function flashAll() {
        document.querySelectorAll('.sensor-card').forEach(c => {
            c.classList.remove('flash');
            void c.offsetWidth;
            c.classList.add('flash');
        });
    }

    function resetSensors() {
        const map = {t1:'°C',t2:'°C',t3:'°C',t4:'°C',t5:'°C',tm:'°C',flow:'m/s',rh:'%',noise:'dB'};
        Object.entries(map).forEach(([k,u]) => setVal(k+'Value','00.00',u));
        currentData = {};
    }

    /* ── RECORDING ────────────────────────────── */
    function parseT(s) {
        const p = s.split(':');
        return p.length === 3 ? (parseInt(p[0])||0)*3600 + (parseInt(p[1])||0)*60 + (parseInt(p[2])||0) : 60;
    }

    function fmtT(sec) {
        const h = Math.floor(sec/3600), m = Math.floor((sec%3600)/60), s = sec%60;
        return [h,m,s].map(n=>String(n).padStart(2,'0')).join(':');
    }

    function startRec() {
        const dur = parseT(document.getElementById('timerInput').value);
        const intv = parseInt(document.getElementById('intervalInput').value) || 2;
        if (dur <= 0) { alert('Enter valid duration.'); return; }
        remaining = dur; isRec = true;
        document.getElementById('saveBtn').classList.add('recording');
        document.getElementById('saveBtnIcon').textContent = '■';
        document.getElementById('saveBtnText').textContent = 'Stop Recording';
        document.getElementById('intervalInput').disabled = true;
        document.getElementById('timerInput').disabled = true;
        timerInt = setInterval(() => {
            remaining--;
            document.getElementById('timerDisplay').textContent = fmtT(remaining);
            if (remaining <= 0) stopRec();
        }, 1000);
        dataInt = setInterval(() => {
            if (Object.keys(currentData).length) addRow(currentData);
        }, intv * 1000);
    }

    function stopRec() {
        isRec = false;
        clearInterval(timerInt); clearInterval(dataInt);
        document.getElementById('saveBtn').classList.remove('recording');
        document.getElementById('saveBtnIcon').textContent = '▶';
        document.getElementById('saveBtnText').textContent = 'Start Recording';
        document.getElementById('intervalInput').disabled = false;
        document.getElementById('timerInput').disabled = false;
    }

    function resetAll() {
        if (!confirm('Reset all recorded data?')) return;
        tableData = [];
        document.getElementById('tableBody').innerHTML = '';
        document.getElementById('rowBadge').textContent = '0 rows';
        if (isRec) stopRec();
        document.getElementById('timerDisplay').textContent = '00:00:00';
    }

    /* ── TABLE ROW ────────────────────────────── */
    function isOOT(d) {
        const t5 = parseFloat(d.t5||0);
        for (const k of ['t1','t2','t3','t4']) {
            const v = parseFloat(d[k]||0);
            if (v>0 && t5>0 && (v<t5-0.8||v>t5+0.8)) return true;
        }
        const rh = parseFloat(d.rh||0);
        if (rh>0 && (rh<40||rh>65)) return true;
        if (parseFloat(d.tm||0)>=40) return true;
        if (parseFloat(d.flow||0)>0.35) return true;
        if (parseFloat(d.noise||0)>=65) return true;
        return false;
    }

    function addRow(d) {
        const now = new Date();
        const row = {
            date: now.toLocaleDateString('id-ID'),
            time: now.toLocaleTimeString('id-ID'),
            t1: +parseFloat(d.t1||0).toFixed(2), t2: +parseFloat(d.t2||0).toFixed(2),
            t3: +parseFloat(d.t3||0).toFixed(2), t4: +parseFloat(d.t4||0).toFixed(2),
            t5: +parseFloat(d.t5||0).toFixed(2), tm: +parseFloat(d.tm||0).toFixed(2),
            rh: +parseFloat(d.rh||0).toFixed(2), flow: +parseFloat(d.flow||0).toFixed(2),
            noise: +parseFloat(d.noise||0).toFixed(2)
        };
        tableData.push(row);
        const tr = document.createElement('tr');
        tr.className = 'new-row' + (isOOT(d) ? ' oot' : '');
        tr.innerHTML = `<td>${row.date}</td><td>${row.time}</td>
            <td>${row.t1}</td><td>${row.t2}</td><td>${row.t3}</td>
            <td>${row.t4}</td><td>${row.t5}</td><td>${row.tm}</td>
            <td>${row.rh}</td><td>${row.flow}</td><td>${row.noise}</td>`;
        const tb = document.getElementById('tableBody');
        tb.insertBefore(tr, tb.firstChild);
        const n = tableData.length;
        document.getElementById('rowBadge').textContent = n + (n===1?' row':' rows');
    }

    /* ── EXPORT ───────────────────────────────── */
    function stats(arr) {
        if (!arr.length) return {min:0,max:0,mean:0,stdev:0};
        const min = Math.min(...arr), max = Math.max(...arr);
        const mean = arr.reduce((a,b)=>a+b,0)/arr.length;
        const stdev = Math.sqrt(arr.reduce((s,v)=>s+Math.pow(v-mean,2),0)/arr.length);
        return {min,max,mean,stdev};
    }

    function exportXLSX() {
        if (!tableData.length) { alert('No data to export.'); return; }
        try {
            const wb = XLSX.utils.book_new();

            // Sheet 1 — Raw Data
            const raw = [['Date','Time','T1 (°C)','T2 (°C)','T3 (°C)','T4 (°C)','T5 (°C)','TM (°C)','Humidity (%)','Airflow (m/s)','Noise (dB)']];
            tableData.forEach(r => raw.push([r.date,r.time,r.t1,r.t2,r.t3,r.t4,r.t5,r.tm,r.rh,r.flow,r.noise]));
            const ws1 = XLSX.utils.aoa_to_sheet(raw);
            ws1['!cols'] = Array(11).fill({wch:14});
            XLSX.utils.book_append_sheet(wb, ws1, 'Raw Data');

            // Sheet 2 — Statistical Analysis
            const st = [['ANALISIS STATISTIK'],[],['Parameter','Minimal','Maksimal','STDEV','Mean'],[]];
            [['T1','t1'],['T2','t2'],['T3','t3'],['T4','t4'],['T5','t5']].forEach(([n,k]) => {
                const v = tableData.map(r=>r[k]).filter(x=>x>0);
                if (v.length) { const s=stats(v); st.push([n,+s.min.toFixed(2),+s.max.toFixed(2),+s.stdev.toFixed(2),+s.mean.toFixed(2)]); }
            });
            st.push([]);
            [['Kelembapan','rh'],['TM (Suhu Matras)','tm'],['Airflow','flow'],['Kebisingan','noise']].forEach(([n,k]) => {
                const v = tableData.map(r=>r[k]).filter(x=>x>0);
                if (v.length) { const s=stats(v); st.push([n,+s.min.toFixed(2),+s.max.toFixed(2),+s.stdev.toFixed(2),+s.mean.toFixed(2)]); }
            });
            const ws2 = XLSX.utils.aoa_to_sheet(st);
            ws2['!cols'] = [{wch:22},{wch:12},{wch:12},{wch:12},{wch:12}];
            XLSX.utils.book_append_sheet(wb, ws2, 'Analisis Statistik');

            const n = new Date();
            XLSX.writeFile(wb, `INCU_${n.getFullYear()}${String(n.getMonth()+1).padStart(2,'0')}${String(n.getDate()).padStart(2,'0')}_${String(n.getHours()).padStart(2,'0')}${String(n.getMinutes()).padStart(2,'0')}.xlsx`);
            alert('Export successful!');
        } catch(e) { alert('Export error: ' + e.message); }
    }

    /* ── EVENTS ───────────────────────────────── */
    document.getElementById('saveBtn').addEventListener('click', () => isRec ? stopRec() : startRec());
    document.getElementById('resetBtn').addEventListener('click', resetAll);
    document.getElementById('exportBtn').addEventListener('click', exportXLSX);
    document.getElementById('timerInput').addEventListener('change', e => {
        if (!isRec) document.getElementById('timerDisplay').textContent = e.target.value || '00:00:00';
    });
    document.getElementById('timerDisplay').textContent = document.getElementById('timerInput').value;
</script>
</body>
</html>
