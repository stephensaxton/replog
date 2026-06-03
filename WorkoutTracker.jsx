import React, { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Plus,
  Upload,
  Download,
  Check,
  Trash2,
  Dumbbell,
  ClipboardList,
  ChevronLeft,
  X,
  RotateCcw,
  CheckCircle2,
  Home,
  Save,
  Library,
} from "lucide-react";

/* ---------- theme ---------- */
const C = {
  bg: "#0c0c0d",
  panel: "#161617",
  panel2: "#1f1f21",
  line: "#2a2a2d",
  text: "#ece8e1",
  dim: "#8a857c",
  accent: "#e0903a", // tin-cloth amber
  accentDim: "#5a4326",
  good: "#5fae6b",
  danger: "#c45c4a",
};

const STORAGE_KEY = "tracker:state";

const uid = () =>
  Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

const today = () => {
  const d = new Date();
  return d.toISOString().slice(0, 10);
};

const prettyDate = (iso) => {
  try {
    return new Date(iso + "T00:00:00").toLocaleDateString(undefined, {
      weekday: "short",
      month: "short",
      day: "numeric",
    });
  } catch {
    return iso;
  }
};

/* ---------- excel helpers ---------- */
const COL_ALIASES = {
  workout: ["workout", "workout name", "routine", "title"],
  exercise: ["exercise", "movement", "lift", "name"],
  sets: ["sets", "set", "# sets", "num sets"],
  reps: ["reps", "rep", "target reps", "rep target", "reps target"],
  weight: ["weight", "load", "lbs", "lb", "kg", "wt"],
};

function findKey(row, aliases) {
  const keys = Object.keys(row);
  for (const k of keys) {
    const norm = String(k).trim().toLowerCase();
    if (aliases.includes(norm)) return k;
  }
  return null;
}

function splitList(raw) {
  // Accept commas, semicolons, slashes, or whitespace as separators. Spaces and
  // slashes survive Excel untouched, whereas a comma-grouped run of 3-digit
  // numbers (e.g. 135,155,175,195) can be coerced into one big number by Excel.
  return raw
    ? String(raw)
        .split(/[,;/\s]+/)
        .map((x) => x.trim())
        .filter((x) => x !== "")
    : [];
}

// For a per-set target list: 1 value applies to every set; a longer list maps
// by index, carrying the last value forward if the list is shorter than sets.
function pickPerSet(list, i) {
  if (list.length === 0) return "";
  if (list.length === 1) return list[0];
  return list[i] ?? list[list.length - 1];
}

function parseRows(rows) {
  if (!rows.length) return [];
  const sample = rows[0];
  const kEx = findKey(sample, COL_ALIASES.exercise);
  const kSets = findKey(sample, COL_ALIASES.sets);
  const kReps = findKey(sample, COL_ALIASES.reps);
  const kWeight = findKey(sample, COL_ALIASES.weight);

  const out = [];
  for (const r of rows) {
    const nameRaw = kEx ? r[kEx] : Object.values(r)[0];
    const name = nameRaw == null ? "" : String(nameRaw).trim();
    if (!name) continue;

    // Reps and Weight may each be a single value (applied to every set) or a
    // comma-separated per-set list, e.g. reps "10,8,5,3" or weight "200,210,220".
    const repList = splitList(kReps ? String(r[kReps] ?? "").trim() : "");
    const weightList = splitList(kWeight ? String(r[kWeight] ?? "").trim() : "");

    // Set count: explicit Sets column wins; otherwise infer from the longer of
    // the two per-set lists; otherwise default to 3.
    const explicitSets = kSets ? parseInt(r[kSets], 10) : NaN;
    const inferred = Math.max(repList.length, weightList.length);
    const setCount = Math.max(
      1,
      Math.min(
        20,
        Number.isFinite(explicitSets) && explicitSets > 0
          ? explicitSets
          : inferred > 1
          ? inferred
          : 3
      )
    );

    const sets = Array.from({ length: setCount }, (_, i) => ({
      id: uid(),
      weight: "",
      targetWeight: pickPerSet(weightList, i),
      reps: "",
      targetReps: pickPerSet(repList, i),
      done: false,
    }));
    out.push({ id: uid(), name, sets });
  }
  return out;
}

function parseWorkoutName(rows) {
  if (!rows.length) return "";
  const key = findKey(rows[0], COL_ALIASES.workout);
  if (!key) return "";
  for (const r of rows) {
    const v = r[key];
    if (v != null && String(v).trim()) return String(v).trim();
  }
  return "";
}

function buildTemplate() {
  const data = [
    { Workout: "Push Day A", Exercise: "Back Squat", Sets: 4, Reps: "10,8,5,3", Weight: "185,205,225,245" },
    { Workout: "", Exercise: "Bench Press", Sets: 3, Reps: 8, Weight: 135 },
    { Workout: "", Exercise: "Bent-Over Row", Sets: 3, Reps: 10, Weight: 95 },
    { Workout: "", Exercise: "Plank (sec)", Sets: 3, Reps: 45, Weight: "" },
  ];
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = [{ wch: 16 }, { wch: 18 }, { wch: 6 }, { wch: 12 }, { wch: 14 }];

  // Force the Reps (col D) and Weight (col E) columns to TEXT so per-set lists
  // like "135,155,175,195" stay intact instead of being read by Excel as one
  // big number. Applied to a block of rows so newly typed entries behave too.
  const TEXT_ROWS = 50;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  range.e.r = Math.max(range.e.r, TEXT_ROWS);
  range.e.c = Math.max(range.e.c, 4);
  ws["!ref"] = XLSX.utils.encode_range(range);
  for (let r = 1; r <= TEXT_ROWS; r++) {
    for (const c of [3, 4]) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell) {
        cell.t = "s";
        cell.v = cell.v == null ? "" : String(cell.v);
        cell.z = "@";
        delete cell.w;
      } else {
        ws[addr] = { t: "s", v: "", z: "@" };
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Workout");
  XLSX.writeFile(wb, "workout-template.xlsx");
}

async function readWorkoutFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  return { name: parseWorkoutName(rows), exercises: parseRows(rows) };
}

/* ---------- storage ---------- */
async function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch {}
  return null;
}
async function saveState(state) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {}
}

/* ---------- small components ---------- */
function NumField({ value, onChange, placeholder, mono = true }) {
  return (
    <input
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      inputMode="decimal"
      style={{
        width: "100%",
        background: C.panel2,
        border: `1px solid ${C.line}`,
        borderRadius: 10,
        color: C.text,
        fontSize: 19,
        fontFamily: mono ? "'JetBrains Mono', monospace" : "inherit",
        textAlign: "center",
        padding: "11px 4px",
        outline: "none",
        WebkitAppearance: "none",
      }}
    />
  );
}

/* ---------- main ---------- */
export default function App() {
  const [state, setState] = useState({ current: null, history: [], library: [] });
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState("home"); // home | today | library | history
  const [viewing, setViewing] = useState(null); // history item being viewed
  const [savedFlash, setSavedFlash] = useState(false); // brief "Saved" confirmation
  const [libFlash, setLibFlash] = useState(false); // brief "Saved to library" confirmation
  const fileRef = useRef(null);
  const libFileRef = useRef(null);

  // load once
  useEffect(() => {
    (async () => {
      const s = await loadState();
      if (s) {
        setState({ current: s.current ?? null, history: s.history ?? [], library: s.library ?? [] });
        if (s.current) setView("today");
      }
      setLoaded(true);
    })();
  }, []);

  // persist on change (after initial load)
  useEffect(() => {
    if (loaded) saveState(state);
  }, [state, loaded]);

  const cur = state.current;

  const setCurrent = useCallback((updater) => {
    setState((s) => ({ ...s, current: updater(s.current) }));
  }, []);

  /* ----- actions ----- */
  const newWorkout = (exercises = [], name = "Workout") =>
    setState((s) => ({
      ...s,
      current: { id: uid(), name, date: today(), exercises },
    }));

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const { name, exercises } = await readWorkoutFile(file);
      if (!exercises.length) {
        alert("No exercises found. Check the column headers (Exercise / Sets / Reps / Weight).");
        return;
      }
      const workoutName = name || file.name.replace(/\.[^.]+$/, "");
      setState((s) => {
        if (s.current) {
          return {
            ...s,
            current: {
              ...s.current,
              exercises: [...s.current.exercises, ...exercises],
            },
          };
        }
        return {
          ...s,
          current: { id: uid(), name: workoutName, date: today(), exercises },
        };
      });
      setView("today");
    } catch (err) {
      alert("Could not read that file. Make sure it's an .xlsx or .csv.");
    }
  };

  const handleLibFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const { name, exercises } = await readWorkoutFile(file);
      if (!exercises.length) {
        alert("No exercises found. Check the column headers (Exercise / Sets / Reps / Weight).");
        return;
      }
      const item = {
        id: uid(),
        name: name || file.name.replace(/\.[^.]+$/, ""),
        date: today(),
        exercises,
      };
      setState((s) => ({ ...s, library: [item, ...s.library] }));
    } catch (err) {
      alert("Could not read that file. Make sure it's an .xlsx or .csv.");
    }
  };

  const addExercise = () =>
    setCurrent((c) => ({
      ...c,
      exercises: [
        ...c.exercises,
        { id: uid(), name: "", sets: [{ id: uid(), weight: "", targetWeight: "", reps: "", targetReps: "", done: false }] },
      ],
    }));

  const updateExName = (exId, name) =>
    setCurrent((c) => ({
      ...c,
      exercises: c.exercises.map((ex) => (ex.id === exId ? { ...ex, name } : ex)),
    }));

  const removeExercise = (exId) =>
    setCurrent((c) => ({ ...c, exercises: c.exercises.filter((ex) => ex.id !== exId) }));

  const addSet = (exId) =>
    setCurrent((c) => ({
      ...c,
      exercises: c.exercises.map((ex) => {
        if (ex.id !== exId) return ex;
        const last = ex.sets[ex.sets.length - 1];
        return {
          ...ex,
          sets: [
            ...ex.sets,
            { id: uid(), weight: "", targetWeight: last?.targetWeight ?? "", reps: "", targetReps: last?.targetReps ?? "", done: false },
          ],
        };
      }),
    }));

  const removeSet = (exId, setId) =>
    setCurrent((c) => ({
      ...c,
      exercises: c.exercises.map((ex) =>
        ex.id === exId ? { ...ex, sets: ex.sets.filter((st) => st.id !== setId) } : ex
      ),
    }));

  const updateSet = (exId, setId, field, val) =>
    setCurrent((c) => ({
      ...c,
      exercises: c.exercises.map((ex) =>
        ex.id === exId
          ? { ...ex, sets: ex.sets.map((st) => (st.id === setId ? { ...st, [field]: val } : st)) }
          : ex
      ),
    }));

  const toggleDone = (exId, setId) =>
    setCurrent((c) => ({
      ...c,
      exercises: c.exercises.map((ex) =>
        ex.id === exId
          ? { ...ex, sets: ex.sets.map((st) => (st.id === setId ? { ...st, done: !st.done } : st)) }
          : ex
      ),
    }));

  const saveWorkout = () => {
    if (!cur) return;
    saveState(state); // immediate write of current progress, no archiving
    setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 1600);
  };

  const finishWorkout = () => {
    if (!cur) return;
    const done = { ...cur, finishedAt: new Date().toISOString() };
    setState((s) => ({ current: null, history: [done, ...s.history] }));
    setView("home");
  };

  const repeat = (item) => {
    if (
      cur &&
      cur.exercises.length &&
      typeof window !== "undefined" &&
      window.confirm &&
      !window.confirm("Start this workout? It will replace the one you have in progress.")
    )
      return;
    const exercises = item.exercises.map((ex) => ({
      id: uid(),
      name: ex.name,
      sets: ex.sets.map((st) => ({
        id: uid(),
        weight: "",
        targetWeight: st.targetWeight || st.weight || "",
        reps: "",
        targetReps: st.targetReps || st.reps || "",
        done: false,
      })),
    }));
    newWorkout(exercises, item.name || "Workout");
    setViewing(null);
    setView("today");
  };

  const startFromLibrary = (item) => repeat(item);

  const saveCurrentToLibrary = () => {
    if (!cur || !cur.exercises.length) return;
    const item = {
      id: uid(),
      name: cur.name || "Workout",
      date: today(),
      exercises: cur.exercises.map((ex) => ({
        id: uid(),
        name: ex.name,
        sets: ex.sets.map((st) => ({
          id: uid(),
          weight: "",
          targetWeight: st.targetWeight || st.weight || "",
          reps: "",
          targetReps: st.targetReps || st.reps || "",
          done: false,
        })),
      })),
    };
    setState((s) => ({ ...s, library: [item, ...s.library] }));
    setLibFlash(true);
    setTimeout(() => setLibFlash(false), 1600);
  };

  const deleteLibrary = (id) =>
    setState((s) => ({ ...s, library: s.library.filter((x) => x.id !== id) }));

  const deleteHistory = (id) =>
    setState((s) => ({ ...s, history: s.history.filter((h) => h.id !== id) }));

  /* ----- derived ----- */
  const totals = (ex) => {
    const all = ex.exercises.flatMap((e) => e.sets);
    return { done: all.filter((s) => s.done).length, total: all.length };
  };

  if (!loaded) {
    return (
      <div style={{ ...page, alignItems: "center", justifyContent: "center" }}>
        <FontStyle />
        <span style={{ color: C.dim }}>Loading…</span>
      </div>
    );
  }

  return (
    <div style={page}>
      <FontStyle />
      <input
        ref={fileRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleFile}
        style={{ display: "none" }}
      />
      <input
        ref={libFileRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleLibFile}
        style={{ display: "none" }}
      />

      {/* header */}
      <header style={header}>
        <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
          <Dumbbell size={20} color={C.accent} />
          <span style={{ fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 700, fontSize: 22, letterSpacing: 1, textTransform: "uppercase" }}>
            Rep Log
          </span>
        </div>
      </header>

      {/* body */}
      <main style={main}>
        {viewing
          ? HistoryDetail({ item: viewing, onBack: () => setViewing(null), onRepeat: () => repeat(viewing) })
          : view === "home"
          ? HomeView()
          : view === "today"
          ? TodayView()
          : view === "library"
          ? LibraryView()
          : HistoryView()}
      </main>

      {/* bottom nav */}
      {!viewing && (
        <nav style={nav}>
          <NavBtn active={view === "home"} onClick={() => setView("home")} icon={<Home size={20} />} label="Home" />
          <NavBtn active={view === "today"} onClick={() => setView("today")} icon={<Dumbbell size={20} />} label="Today" />
          <NavBtn active={view === "library"} onClick={() => setView("library")} icon={<Library size={20} />} label="Library" />
          <NavBtn active={view === "history"} onClick={() => setView("history")} icon={<ClipboardList size={20} />} label="History" />
        </nav>
      )}
    </div>
  );

  /* ---------- inner views ---------- */
  function HomeView() {
    return (
      <div style={{ padding: "32px 18px", textAlign: "center", color: C.dim }}>
        <Dumbbell size={44} color={C.line} style={{ marginBottom: 14 }} />
        <p style={{ fontSize: 16, margin: "0 0 26px" }}>
          {cur ? "Workout in progress." : "Ready to train."}
        </p>
        <div style={{ display: "flex", flexDirection: "column", gap: 12, maxWidth: 320, margin: "0 auto" }}>
          {cur && (
            <BigBtn primary onClick={() => setView("today")} icon={<Dumbbell size={18} />}>
              Resume current workout
            </BigBtn>
          )}
          <BigBtn primary={!cur} onClick={() => fileRef.current?.click()} icon={<Upload size={18} />}>
            Import from Excel
          </BigBtn>
          <BigBtn onClick={() => { newWorkout([]); setView("today"); }} icon={<Plus size={18} />}>
            Start blank workout
          </BigBtn>
          <BigBtn onClick={() => setView("library")} icon={<Library size={18} />}>
            Open library
          </BigBtn>
          <BigBtn onClick={buildTemplate} icon={<Download size={18} />}>
            Download template
          </BigBtn>
          {state.history[0] && (
            <BigBtn onClick={() => repeat(state.history[0])} icon={<RotateCcw size={18} />}>
              Repeat last workout
            </BigBtn>
          )}
        </div>
      </div>
    );
  }

  function TodayView() {
    if (!cur) {
      return (
        <div style={{ padding: "40px 18px", textAlign: "center", color: C.dim }}>
          <Dumbbell size={44} color={C.line} style={{ marginBottom: 14 }} />
          <p style={{ fontSize: 16, margin: "0 0 22px" }}>No workout in progress.</p>
          <div style={{ maxWidth: 320, margin: "0 auto" }}>
            <BigBtn primary onClick={() => setView("home")} icon={<Home size={18} />}>
              Go to Home
            </BigBtn>
          </div>
        </div>
      );
    }

    const t = totals(cur);
    return (
      <div style={{ padding: "16px 14px 30px" }}>
        {/* workout meta */}
        <div style={{ marginBottom: 16 }}>
          <input
            value={cur.name}
            onChange={(e) => setCurrent((c) => ({ ...c, name: e.target.value }))}
            style={{
              background: "transparent",
              border: "none",
              borderBottom: `1px solid ${C.line}`,
              color: C.text,
              fontFamily: "'Barlow Condensed', sans-serif",
              fontWeight: 700,
              fontSize: 30,
              letterSpacing: 0.5,
              width: "100%",
              padding: "2px 0 6px",
              outline: "none",
            }}
          />
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
            <span style={{ color: C.dim, fontSize: 13 }}>{prettyDate(cur.date)}</span>
            <span style={{ color: t.done === t.total && t.total > 0 ? C.good : C.dim, fontSize: 13, fontFamily: "'JetBrains Mono', monospace" }}>
              {t.done}/{t.total} sets
            </span>
          </div>
        </div>

        {/* exercises */}
        {cur.exercises.map((ex, i) => (
          <React.Fragment key={ex.id}>{ExerciseCard({ ex, index: i })}</React.Fragment>
        ))}

        {/* add row */}
        <div style={{ display: "flex", gap: 10, marginTop: 6 }}>
          <BigBtn onClick={addExercise} icon={<Plus size={18} />} style={{ flex: 1 }}>
            Exercise
          </BigBtn>
          <BigBtn onClick={() => fileRef.current?.click()} icon={<Upload size={18} />} style={{ flex: 1 }}>
            Import
          </BigBtn>
        </div>

        {cur.exercises.length > 0 && (
          <BigBtn
            onClick={saveCurrentToLibrary}
            icon={libFlash ? <Check size={18} /> : <Library size={18} />}
            style={{ marginTop: 10, color: libFlash ? C.good : C.text, borderColor: libFlash ? C.good : C.line }}
          >
            {libFlash ? "Saved to library" : "Save to library"}
          </BigBtn>
        )}

        {cur.exercises.length > 0 && (
          <div style={{ display: "flex", gap: 10, marginTop: 18 }}>
            <BigBtn primary onClick={finishWorkout} icon={<CheckCircle2 size={18} />} style={{ flex: 1 }}>
              Finish workout
            </BigBtn>
            <BigBtn
              onClick={saveWorkout}
              icon={savedFlash ? <Check size={18} /> : <Save size={18} />}
              style={{ flex: 1, color: savedFlash ? C.good : C.text, borderColor: savedFlash ? C.good : C.line }}
            >
              {savedFlash ? "Saved" : "Save progress"}
            </BigBtn>
          </div>
        )}
      </div>
    );
  }

  function ExerciseCard({ ex, index }) {
    return (
      <div style={card}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
          <span style={{ color: C.accent, fontFamily: "'JetBrains Mono', monospace", fontSize: 13, minWidth: 18 }}>
            {String(index + 1).padStart(2, "0")}
          </span>
          <input
            value={ex.name}
            onChange={(e) => updateExName(ex.id, e.target.value)}
            placeholder="Exercise name"
            style={{
              flex: 1,
              background: "transparent",
              border: "none",
              color: C.text,
              fontSize: 17,
              fontWeight: 600,
              outline: "none",
            }}
          />
          <button onClick={() => removeExercise(ex.id)} style={iconBtn} aria-label="Remove exercise">
            <Trash2 size={16} color={C.dim} />
          </button>
        </div>

        {/* column labels */}
        <div style={{ display: "grid", gridTemplateColumns: "26px 1fr 1fr 44px", gap: 8, marginBottom: 6, padding: "0 2px" }}>
          <span style={lbl}>#</span>
          <span style={{ ...lbl, textAlign: "center" }}>Weight</span>
          <span style={{ ...lbl, textAlign: "center" }}>Reps</span>
          <span />
        </div>

        {ex.sets.map((st, si) => (
          <div
            key={st.id}
            style={{
              display: "grid",
              gridTemplateColumns: "26px 1fr 1fr 44px",
              gap: 8,
              alignItems: "center",
              marginBottom: 8,
              opacity: st.done ? 0.55 : 1,
            }}
          >
            <span style={{ color: C.dim, fontFamily: "'JetBrains Mono', monospace", fontSize: 14, textAlign: "center" }}>
              {si + 1}
            </span>
            <NumField
              value={st.weight}
              onChange={(v) => updateSet(ex.id, st.id, "weight", v)}
              placeholder={st.targetWeight ? String(st.targetWeight) : "—"}
            />
            <NumField
              value={st.reps}
              onChange={(v) => updateSet(ex.id, st.id, "reps", v)}
              placeholder={st.targetReps ? `×${st.targetReps}` : "—"}
            />
            <div style={{ display: "flex", gap: 2 }}>
              <button
                onClick={() => toggleDone(ex.id, st.id)}
                style={{
                  ...iconBtn,
                  width: 38,
                  height: 38,
                  borderRadius: 10,
                  background: st.done ? C.good : C.panel2,
                  border: `1px solid ${st.done ? C.good : C.line}`,
                }}
                aria-label="Mark set done"
              >
                <Check size={18} color={st.done ? "#0c0c0d" : C.dim} strokeWidth={3} />
              </button>
            </div>
          </div>
        ))}

        <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
          <button onClick={() => addSet(ex.id)} style={ghostBtn}>
            <Plus size={14} /> Add set
          </button>
          {ex.sets.length > 1 && (
            <button onClick={() => removeSet(ex.id, ex.sets[ex.sets.length - 1].id)} style={ghostBtn}>
              <X size={14} /> Remove set
            </button>
          )}
        </div>
      </div>
    );
  }

  function LibraryView() {
    return (
      <div style={{ padding: "16px 14px 30px" }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 18 }}>
          <BigBtn primary onClick={() => libFileRef.current?.click()} icon={<Upload size={18} />}>
            Import workout to library
          </BigBtn>
          <BigBtn onClick={buildTemplate} icon={<Download size={18} />}>
            Download template
          </BigBtn>
        </div>

        {state.library.length === 0 ? (
          <div style={{ padding: "30px 10px", textAlign: "center", color: C.dim }}>
            <Library size={40} color={C.line} style={{ marginBottom: 12 }} />
            <p style={{ margin: 0 }}>No saved workouts in your library yet.</p>
            <p style={{ margin: "8px 0 0", fontSize: 13 }}>
              Import an Excel workout above to keep it on hand.
            </p>
          </div>
        ) : (
          state.library.map((item) => (
            <div key={item.id} style={card}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 17, fontWeight: 600 }}>{item.name || "Workout"}</div>
                  <div style={{ color: C.dim, fontSize: 13, marginTop: 4, fontFamily: "'JetBrains Mono', monospace" }}>
                    {item.exercises.length} exercises
                  </div>
                  <div
                    style={{
                      color: C.dim,
                      fontSize: 12,
                      marginTop: 6,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {item.exercises.map((e) => e.name).filter(Boolean).join(" · ") || "—"}
                  </div>
                </div>
                <button onClick={() => deleteLibrary(item.id)} style={iconBtn} aria-label="Delete from library">
                  <Trash2 size={16} color={C.dim} />
                </button>
              </div>
              <BigBtn primary onClick={() => startFromLibrary(item)} icon={<Dumbbell size={18} />} style={{ marginTop: 12 }}>
                Start workout
              </BigBtn>
            </div>
          ))
        )}
      </div>
    );
  }

  function HistoryView() {
    if (!state.history.length) {
      return (
        <div style={{ padding: "50px 18px", textAlign: "center", color: C.dim }}>
          <ClipboardList size={40} color={C.line} style={{ marginBottom: 12 }} />
          <p>No saved workouts yet.</p>
        </div>
      );
    }
    return (
      <div style={{ padding: "16px 14px 30px" }}>
        {state.history.map((h) => {
          const t = totals(h);
          return (
            <button key={h.id} onClick={() => setViewing(h)} style={{ ...card, width: "100%", textAlign: "left", cursor: "pointer", display: "block" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
                <span style={{ fontSize: 17, fontWeight: 600 }}>{h.name || "Workout"}</span>
                <span style={{ color: C.dim, fontSize: 12 }}>{prettyDate(h.date)}</span>
              </div>
              <div style={{ color: C.dim, fontSize: 13, marginTop: 6, fontFamily: "'JetBrains Mono', monospace" }}>
                {h.exercises.length} exercises · {t.done}/{t.total} sets
              </div>
            </button>
          );
        })}
      </div>
    );
  }

  function HistoryDetail({ item, onBack, onRepeat }) {
    return (
      <div style={{ padding: "14px 14px 30px" }}>
        <button onClick={onBack} style={{ ...ghostBtn, marginBottom: 14 }}>
          <ChevronLeft size={16} /> Back
        </button>
        <h2 style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 28, fontWeight: 700, margin: "0 0 2px" }}>
          {item.name || "Workout"}
        </h2>
        <p style={{ color: C.dim, fontSize: 13, margin: "0 0 16px" }}>{prettyDate(item.date)}</p>

        {item.exercises.map((ex) => (
          <div key={ex.id} style={card}>
            <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>{ex.name || "—"}</div>
            {ex.sets.map((st, i) => (
              <div key={st.id} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: i < ex.sets.length - 1 ? `1px solid ${C.line}` : "none", fontFamily: "'JetBrains Mono', monospace", fontSize: 14 }}>
                <span style={{ color: C.dim }}>Set {i + 1}</span>
                <span style={{ color: st.done ? C.good : C.text }}>
                  {st.weight ? `${st.weight}` : "—"} × {st.reps ? st.reps : "—"}
                </span>
              </div>
            ))}
          </div>
        ))}

        <div style={{ display: "flex", gap: 10, marginTop: 8 }}>
          <BigBtn primary onClick={onRepeat} icon={<RotateCcw size={18} />} style={{ flex: 1 }}>
            Repeat
          </BigBtn>
          <BigBtn onClick={() => { deleteHistory(item.id); onBack(); }} icon={<Trash2 size={18} />} style={{ flex: 1 }}>
            Delete
          </BigBtn>
        </div>
      </div>
    );
  }
}

/* ---------- reusable bits ---------- */
function NavBtn({ active, onClick, icon, label }) {
  return (
    <button onClick={onClick} style={{ flex: 1, background: "transparent", border: "none", color: active ? C.accent : C.dim, display: "flex", flexDirection: "column", alignItems: "center", gap: 3, padding: "8px 0", cursor: "pointer" }}>
      {icon}
      <span style={{ fontSize: 11, letterSpacing: 0.5 }}>{label}</span>
    </button>
  );
}

function BigBtn({ children, onClick, icon, primary, style }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: 8,
        background: primary ? C.accent : C.panel,
        color: primary ? "#0c0c0d" : C.text,
        border: `1px solid ${primary ? C.accent : C.line}`,
        borderRadius: 12,
        padding: "13px 14px",
        fontSize: 15,
        fontWeight: 600,
        cursor: "pointer",
        ...style,
      }}
    >
      {icon}
      {children}
    </button>
  );
}

/* ---------- styles ---------- */
const page = {
  display: "flex",
  flexDirection: "column",
  height: "100%",
  minHeight: 560,
  maxWidth: 480,
  margin: "0 auto",
  background: C.bg,
  color: C.text,
  fontFamily: "'Barlow', system-ui, sans-serif",
};
const header = {
  padding: "14px 16px",
  borderBottom: `1px solid ${C.line}`,
  position: "sticky",
  top: 0,
  background: C.bg,
  zIndex: 5,
};
const main = { flex: 1, overflowY: "auto" };
const nav = {
  display: "flex",
  borderTop: `1px solid ${C.line}`,
  background: C.bg,
  position: "sticky",
  bottom: 0,
};
const card = {
  background: C.panel,
  border: `1px solid ${C.line}`,
  borderRadius: 14,
  padding: 14,
  marginBottom: 12,
};
const lbl = { color: C.dim, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.6 };
const iconBtn = {
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  background: "transparent",
  border: "none",
  cursor: "pointer",
  padding: 4,
};
const ghostBtn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 5,
  background: "transparent",
  border: `1px solid ${C.line}`,
  borderRadius: 9,
  color: C.dim,
  fontSize: 13,
  padding: "7px 11px",
  cursor: "pointer",
};

function FontStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;600&family=Barlow+Condensed:wght@600;700&family=JetBrains+Mono:wght@400;500&display=swap');
      * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
      input::placeholder { color: ${C.dim}; opacity: 0.7; }
      ::-webkit-scrollbar { width: 0; }
    `}</style>
  );
}
